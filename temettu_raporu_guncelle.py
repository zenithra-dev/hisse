"""
Temettü Raporu güncelleyici.

Girdi : kaynak/temettu_raporu_orijinal.xlsx  (BIST ve ABD ilk 20 temettü raporu)
        temettu_fonlari.json                 (temettü dağıtan ABD fonları / ETF verisi)
        temettu_cin.json                     (Çin'in temettü veren en büyük 20 şirketi)
        temettu_cin_fonlari.json             (Çin temettü fonları / ETF)
Çıktı : temettu_raporu.xlsx      (BIST + ABD + ABD temettü fonları)
        temettu_raporu_cin.xlsx  (Çin - ayrı dosya)

Yaptığı değişiklikler:
1. "Son 3/5/7/10 yıl kesintisiz" sütunları Evet/Hayır yerine SAYI gösterir:
   10 yılın 9'unda temettü ödediyse 9 yazar. Hepsinde ödediyse yeşil, eksik varsa pembe.
2. Özet, "Batmama Riski" sayfası ve alt toplamlar yeni sayılara göre çalışır.
3. "ABD Temettü Fonları" sayfası eklenir (aynı ✓/✗ ve sayı mantığıyla).
4. Çin için AYRI bir dosya üretilir: "Çin İlk 20" (Hong Kong / ABD'de işlem gören
   Çin şirketleri), "Çin Temettü Fonları", "Batmama Riski (3 yıl)" ve "Not Tablosu".

Kullanım:  python3 temettu_raporu_guncelle.py
"""
import json
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

KOK = Path(__file__).resolve().parent
GIRDI = KOK / "kaynak" / "temettu_raporu_orijinal.xlsx"
FON_VERISI = KOK / "temettu_fonlari.json"
CIN_VERISI = KOK / "temettu_cin.json"
CIN_FON_VERISI = KOK / "temettu_cin_fonlari.json"
CIKTI = KOK / "temettu_raporu.xlsx"
CIN_CIKTI = KOK / "temettu_raporu_cin.xlsx"

ILK, SON = 5, 24  # şirket satırları
YIL_SUTUN = "FGHIJKLMNO"  # 2017 .. 2026
# (sütun, yıl sayısı, ilk yıl sütunu)
SAYAC = [("P", 3, "M"), ("Q", 5, "K"), ("R", 7, "I"), ("S", 10, "F")]

YESIL = PatternFill("solid", fgColor="FFC6EFCE")
PEMBE = PatternFill("solid", fgColor="FFFFC7CE")
GRI = PatternFill("solid", fgColor="FFEDEDED")
LACIVERT = PatternFill("solid", fgColor="FF1F3864")
ACIK_MAVI = PatternFill("solid", fgColor="FFD9E1F2")
INCE = Side(style="thin", color="FFBFBFBF")
KENAR = Border(left=INCE, right=INCE, top=INCE, bottom=INCE)


def baslik(n):
    return f"Son {n} yılda\nkaç yıl\ntemettü verdi"


def sayac_formulu(sutun_ilk, satir):
    return f"=COUNTIF({sutun_ilk}{satir}:O{satir},1)"


def geriye_kesintisiz(r):
    """2026'dan geriye kaç yıl üst üste ödendi (en çok 10)."""
    f = "10"
    for k, col in enumerate(YIL_SUTUN):  # F (2017) .. O (2026)
        f = f"IF(N({col}{r})=0,{9 - k},{f})"
    return "=" + f


def sayac_renkleri(ws, ilk=ILK, son=SON):
    """Sayı sütunlarının eski Evet/Hayır renklerini kaldırıp yenilerini ekler."""
    eski = ws.conditional_formatting
    yeni = type(eski)()
    for aralik in eski:
        if str(aralik.sqref).startswith("P"):
            continue  # Evet/Hayır kuralları
        for kural in aralik.rules:
            yeni.add(str(aralik.sqref), kural)
    ws.conditional_formatting = yeni
    for sutun, n, _ in SAYAC:
        rng = f"{sutun}{ilk}:{sutun}{son}"
        hucre = f"{sutun}{ilk}"
        ws.conditional_formatting.add(
            rng, FormulaRule(formula=[f"AND(ISNUMBER({hucre}),{hucre}={n})"], fill=YESIL))
        ws.conditional_formatting.add(
            rng, FormulaRule(formula=[f"AND(ISNUMBER({hucre}),{hucre}<{n})"], fill=PEMBE))


def sirket_sayfasi(ws):
    for sutun, n, ilk_yil in SAYAC:
        ws[f"{sutun}4"] = baslik(n)
        for r in range(ILK, SON + 1):
            c = ws[f"{sutun}{r}"]
            c.value = sayac_formulu(ilk_yil, r)
            c.number_format = "0"
        ws[f"{sutun}26"] = f"=COUNTIF({sutun}{ILK}:{sutun}{SON},{n})"
    ws["C26"] = "Hiç aksatmadan ödeyen şirket sayısı:"
    ws["C27"] = ("✓ = o yıl nakit temettü ödendi, ✗ = ödenmedi, gri = veri yok / halka arz öncesi "
                 "(sayılmaz). Sayı sütunları: o dönemde kaç yıl temettü ödendiği "
                 "(ör. 10 yılda 9 = 9). Mavi yazı = elle girilen veri, siyah = formül.")
    for col in "PQRS":
        ws.column_dimensions[col].width = 11
    ws.row_dimensions[4].height = 45
    sayac_renkleri(ws)


def stil_kopyala(hedef, kaynak, alanlar=("font", "fill", "border", "alignment", "number_format")):
    for attr in alanlar:
        setattr(hedef, attr, copy(getattr(kaynak, attr)))


def ozet_sayfasi(ws, fon_sayisi):
    """Sütunlar: A etiket | BIST | ABD | Toplam (şirket) | (ABD temettü fonları)."""
    veri_stili, toplam_stili, baslik_stili = {}, {}, {}
    for attr in ("font", "fill", "border", "alignment", "number_format"):
        veri_stili[attr] = copy(getattr(ws["C5"], attr))
        toplam_stili[attr] = copy(getattr(ws["D5"], attr))
        baslik_stili[attr] = copy(getattr(ws["C4"], attr))
    baslik_stili["alignment"] = Alignment(horizontal="center", vertical="center", wrap_text=True)

    sayfalar = [("BIST (20 şirket)", "BIST İlk 20", 24), ("ABD (20 şirket)", "ABD İlk 20", 24)]
    sirket_sutunlari = [chr(ord("B") + i) for i in range(len(sayfalar))]
    toplam_sutun = chr(ord("B") + len(sayfalar))
    fon_sutun = chr(ord(toplam_sutun) + 1)

    def yaz(hucre, deger, stil):
        c = ws[hucre]
        c.value = deger
        for attr, v in stil.items():
            setattr(c, attr, copy(v))

    ws["A4"] = "Hiç aksatmadan temettü ödeyen"
    for col, (ad, _, _) in zip(sirket_sutunlari, sayfalar):
        yaz(f"{col}4", ad, baslik_stili)
    yaz(f"{toplam_sutun}4", "Toplam\n(şirket)", baslik_stili)
    if fon_sayisi:
        yaz(f"{fon_sutun}4", f"ABD temettü fonları\n({fon_sayisi} fon)", baslik_stili)

    for r, (s, n) in zip(range(5, 9), [("P", 3), ("Q", 5), ("R", 7), ("S", 10)]):
        ws[f"A{r}"] = f"Son {n} yılın hepsinde ödeyen"
        for col, (_, sayfa, son) in zip(sirket_sutunlari, sayfalar):
            yaz(f"{col}{r}", f"=COUNTIF('{sayfa}'!{s}5:{s}{son},{n})", veri_stili)
        yaz(f"{toplam_sutun}{r}", f"=SUM({sirket_sutunlari[0]}{r}:{sirket_sutunlari[-1]}{r})", toplam_stili)
        if fon_sayisi:
            yaz(f"{fon_sutun}{r}",
                f"=COUNTIF('ABD Temettü Fonları'!{s}5:{s}{4 + fon_sayisi},{n})", veri_stili)

    ws.row_dimensions[4].height = 30
    for col in sirket_sutunlari + [toplam_sutun]:
        ws.column_dimensions[col].width = 18
    ws.column_dimensions[fon_sutun].width = 20

    ws["A12"] = ("• BIST İlk 20 / ABD İlk 20: her yıl temettü ödendi mi (✓/✗), son 3-5-7-10 yılda "
                 "kaç yıl ödendiği (ör. 10 yılda 9), verim, kredi notu, borçluluk.")
    if fon_sayisi:
        ws["A13"] = ("• ABD Temettü Fonları: temettü dağıtan en büyük 20 ABD borsa yatırım fonu (ETF); "
                     "aynı ✓/✗ ve sayı mantığı, gider oranı ve verim.")
    ws["A14"] = ("• Batmama Riski (3 yıl): iki borsadaki 40 şirket için basit finansal sağlamlık "
                 "göstergesi.  • Not Tablosu: kredi notlarının kategori karşılıkları.")
    ws["A17"] = ("Yıl kolonlarındaki ✓/✗ hücreleri düzeltilirse tüm sayılar ve göstergeler "
                 "otomatik güncellenir.")


def gosterge_formulu(r):
    return (f'=IF(AND(E{r}="Yatırım yapılabilir",F{r}=10),"Çok güçlü",'
            f'IF(OR(E{r}="Yatırım yapılabilir",F{r}=10),"Güçlü",'
            f'IF(G{r}=5,"Orta","Takip edilmeli")))')


def batmama_sayfasi(ws):
    ws["F4"] = "Son 10 yılda\nkaç yıl temettü"
    ws["G4"] = "Son 5 yılda\nkaç yıl temettü"
    for r in range(5, 45):
        ws[f"F{r}"].number_format = "0"
        ws[f"G{r}"].number_format = "0"
        ws[f"J{r}"] = gosterge_formulu(r)
    ws["A47"] = ("• Çok güçlü: yatırım yapılabilir kredi notu (BBB- / Baa3 ve üstü) VE son 10 yılın "
                 "10'unda da temettü.")
    ws["A49"] = "• Orta: yalnızca son 5 yılın 5'inde de temettü."


def cin_batmama_sayfasi(ws, n):
    """Batmama sayfasını yalnızca Çin şirketlerini (n adet) gösterecek şekilde yeniden kurar.

    Önce batmama_sayfasi() çalıştırılmış olmalı. Satır 5..4+n Çin'e bağlanır, kalan
    şirket satırları silinir, alttaki açıklama bloğu (46-59) yukarı taşınır.
    """
    son = 4 + n
    for i in range(n):
        r = 5 + i
        degerler = {
            "A": "Çin",
            "B": f"='Çin İlk 20'!B{r}",
            "C": f"='Çin İlk 20'!C{r}",
            "D": f"='Çin İlk 20'!X{r}",
            "F": f"='Çin İlk 20'!S{r}",
            "G": f"='Çin İlk 20'!Q{r}",
            "H": f"='Çin İlk 20'!Z{r}&\"\"",
            "I": f"=IF('Çin İlk 20'!AA{r}=\"\",\"\",'Çin İlk 20'!AA{r})",
        }
        for col, v in degerler.items():
            ws[f"{col}{r}"] = v
    for r in range(son + 1, 45):  # kullanılmayan şirket satırlarını tamamen kaldır
        for col in range(1, 11):
            ws._cells.pop((r, col), None)
    kaydir = 44 - son
    if kaydir:
        ws.move_range("A46:J59", rows=-kaydir)
    t = lambda r: r - kaydir  # eski satır -> yeni satır
    ws[f"A{t(48)}"] = "• Güçlü: bu ikisinden yalnızca biri (ör. yüksek not ama temettü geçmişi kısa: Alibaba)."
    ws[f"A{t(51)}"] = ("Kredi notu 'Doğrulanamadı' yazan şirketler için güncel not kaynağı bulunamadı; burada "
                       "'Not yok' sayıldıkları için göstergeleri olduğundan düşük çıkmış olabilir.")
    ws[f"A{t(52)}"] = ("Çin'de kurulu şirketlerin (H hissesi) temettüsünden %10 stopaj kesilir; ayrıntı "
                       "'Çin İlk 20' sayfasının altındaki vergi notunda.")
    for r in range(t(56), t(60)):
        ws[f"C{r}"] = f"=COUNTIF(J5:J{son},A{r})"
    eski = ws.conditional_formatting
    yeni = type(eski)()
    for aralik in eski:
        for kural in aralik.rules:
            yeni.add(f"J5:J{son}", kural)
    ws.conditional_formatting = yeni
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["D"].width = 15


# ---------------------------------------------------------------- Çin sayfası
CIN_BASLIK = {
    "A": "Sıra", "B": "Kod\n(Hong Kong)", "C": "Şirket", "D": "Sektör",
    "E": "Piyasa Değeri\n(milyar USD)",
    "T": "Geriye doğru\nkesintisiz yıl\n(tabloda, max 10)",
    "U": "2026 hisse\nbaşı temettü\n(brüt)", "V": "Para\nbirimi", "W": "Temettü\nverimi",
    "X": "Kredi notu", "Y": "Not veren", "Z": "Borçluluk ölçüsü", "AA": "Değer",
    "AB": "ABD kodu\n(ADR)", "AC": "Lot\n(en az adet)", "AD": "Not / Kaynak",
}
CIN_GENISLIK = {"A": 5, "B": 10, "C": 28, "D": 20, "E": 12, "T": 12, "U": 12, "V": 8, "W": 9,
                "X": 10, "Y": 11, "Z": 20, "AA": 9, "AB": 10, "AC": 10, "AD": 60}


def cin_sayfasi(wb, veri):
    sablon = wb["BIST İlk 20"]
    ws = wb.create_sheet("Çin İlk 20", index=0)
    sirketler = veri["companies"]
    son = 4 + len(sirketler)
    stil = stil_kopyala

    ws["A1"] = "Çin — son 3 yılda temettü veren en büyük 20 şirket (Hong Kong / ABD'de işlem görenler)"
    stil(ws["A1"], sablon["A1"])
    ws["A2"] = (f"Veri tarihi: {veri['as_of']}. Yalnızca Türkiye'den alınabilen (Hong Kong borsası veya "
                "ABD'de ADR) Çin şirketleri. Seçim kriteri ve elenen şirketler tablonun altında.")
    stil(ws["A2"], sablon["A2"])

    for col, metin in CIN_BASLIK.items():
        ws[f"{col}4"] = metin
    for i, col in enumerate(YIL_SUTUN):
        ws[f"{col}4"] = 2017 + i
    for sutun, n, _ in SAYAC:
        ws[f"{sutun}4"] = baslik(n)
    for col in list(CIN_BASLIK) + list(YIL_SUTUN) + [x for x, _, _ in SAYAC]:
        stil(ws[f"{col}4"], sablon["A4"])
    ws.row_dimensions[4].height = 45

    for i, f in enumerate(sirketler):
        r = ILK + i
        kod = f"{f['hk_code']}.HK" if f.get("hk_code") else (f.get("us_ticker") or "")
        degerler = {
            "A": i + 1, "B": kod, "C": f["name"], "D": f.get("sektor") or f.get("sector"),
            "E": f.get("mcap_bn_usd"), "U": f.get("last_dps"), "V": f.get("dps_currency"),
            "W": f.get("yield"), "X": f.get("rating") or "Yok", "Y": f.get("rating_agency"),
            "Z": f.get("borcluluk") or f.get("leverage_measure"), "AA": f.get("leverage_value"),
            "AB": f.get("us_ticker") if f.get("hk_code") else None, "AC": f.get("board_lot"),
            "AD": " | ".join(x for x in [f.get("not") or f.get("note"),
                                         "Kaynak: " + f["source_url"] if f.get("source_url") else None] if x),
        }
        for col, v in degerler.items():
            ws[f"{col}{r}"] = v
        # stiller: BIST İlk 20 satır 5'ten
        for col, kaynak in {"A": "A", "B": "B", "C": "C", "D": "D", "E": "E", "T": "T", "U": "U",
                            "V": "W", "W": "V", "X": "W", "Y": "X", "Z": "Y", "AA": "Z",
                            "AB": "W", "AC": "W", "AD": "AA"}.items():
            stil(ws[f"{col}{r}"], sablon[f"{kaynak}5"])
        ws[f"U{r}"].number_format = "#,##0.000"
        ws[f"AC{r}"].number_format = "#,##0"
        ws[f"AD{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        for j, col in enumerate(YIL_SUTUN):
            c = ws[f"{col}{r}"]
            c.value = f["years"].get(str(2017 + j))
            stil(c, sablon["F5"])
            if c.value is None:
                c.fill = copy(GRI)
        for sutun, n, ilk_yil in SAYAC:
            c = ws[f"{sutun}{r}"]
            c.value = sayac_formulu(ilk_yil, r)
            stil(c, sablon["P5"])
            c.number_format = "0"
        ws[f"T{r}"] = geriye_kesintisiz(r)
        ws.row_dimensions[r].height = 40

    ws.conditional_formatting.add(
        f"F{ILK}:O{son}", FormulaRule(formula=["AND(ISNUMBER(F5),F5=1)"], fill=YESIL))
    ws.conditional_formatting.add(
        f"F{ILK}:O{son}", FormulaRule(formula=["AND(ISNUMBER(F5),F5=0)"], fill=PEMBE))
    sayac_renkleri(ws, ILK, son)

    alt = son + 2
    ws[f"C{alt}"] = "Hiç aksatmadan ödeyen şirket sayısı:"
    stil(ws[f"C{alt}"], sablon["C26"])
    for sutun, n, _ in SAYAC:
        ws[f"{sutun}{alt}"] = f"=COUNTIF({sutun}{ILK}:{sutun}{son},{n})"
        stil(ws[f"{sutun}{alt}"], sablon["P26"])
    ws[f"C{alt + 1}"] = ("✓ = o yıl nakit temettü ödendi, ✗ = ödenmedi, gri = şirket henüz halka açık "
                         "değildi (sayılmaz). Sayı sütunları: o dönemde kaç yıl temettü ödendiği. "
                         "Mavi yazı = elle girilen veri, siyah = formül.")
    stil(ws[f"C{alt + 1}"], sablon["C27"])

    satir = alt + 3
    for metin, yukseklik in [(veri.get("selection_note"), 110), (veri.get("tax_note"), 70)]:
        if not metin:
            continue
        ws[f"A{satir}"] = metin
        stil(ws[f"A{satir}"], sablon["A29"])
        ws.merge_cells(f"A{satir}:T{satir}")
        ws.row_dimensions[satir].height = yukseklik
        satir += 1

    for col, w in CIN_GENISLIK.items():
        ws.column_dimensions[col].width = w
    for col in YIL_SUTUN:
        ws.column_dimensions[col].width = 6
    for col in "PQRS":
        ws.column_dimensions[col].width = 11
    ws.freeze_panes = "D5"
    ws.sheet_view.zoomScale = sablon.sheet_view.zoomScale
    return len(sirketler)


# ---------------------------------------------------------------- fon sayfası
FON_BASLIK = {
    "A": "Sıra", "B": "Kod", "C": "Fon", "D": "Ne yapar (strateji)",
    "E": "Fon büyüklüğü\n(milyar USD)",
    "T": "Geriye doğru\nkesintisiz yıl\n(tabloda, max 10)",
    "U": "İhraççı", "V": "Kuruluş", "W": "Ödeme\nsıklığı",
    "X": "Son 12 ay\ntemettü / pay\n(USD)", "Y": "Temettü\nverimi", "Z": "Yıllık gider\noranı",
    "AA": "Not / Kaynak",
}
FON_GENISLIK = {"A": 5, "B": 8, "C": 30, "D": 32, "E": 12, "T": 12, "U": 14, "V": 11,
                "W": 10, "X": 11, "Y": 9, "Z": 10, "AA": 60}
SIKLIK = {"monthly": "Aylık", "quarterly": "3 ayda bir", "annual": "Yılda bir",
          "semiannual": "6 ayda bir", "semi-annual": "6 ayda bir"}


def tarih_tr(iso):
    """2011-10-20 -> 20.10.2011"""
    if not iso or len(iso) != 10:
        return iso
    y, m, d = iso.split("-")
    return f"{d}.{m}.{y}"


def fon_sayfasi(wb, veri):
    sablon = wb["ABD İlk 20"]
    ws = wb.create_sheet("ABD Temettü Fonları", index=wb.sheetnames.index("ABD İlk 20") + 1)
    fonlar = veri["funds"]
    son = 4 + len(fonlar)

    def stil(hedef, kaynak):
        for attr in ("font", "fill", "border", "alignment", "number_format"):
            setattr(hedef, attr, copy(getattr(kaynak, attr)))

    ws["A1"] = "ABD — temettü dağıtan en büyük 20 borsa yatırım fonu (ETF)"
    stil(ws["A1"], sablon["A1"])
    ws["A2"] = (f"Veri tarihi: {veri['as_of']}. Fon = içinde onlarca/yüzlerce hisse olan sepet; "
                "tek şirkete göre riski dağıtır. Seçim kriteri ve elenen fonlar tablonun altında.")
    stil(ws["A2"], sablon["A2"])

    for col, metin in FON_BASLIK.items():
        ws[f"{col}4"] = metin
    for i, col in enumerate(YIL_SUTUN):
        ws[f"{col}4"] = 2017 + i
    for sutun, n, _ in SAYAC:
        ws[f"{sutun}4"] = baslik(n)
    for col in list(FON_BASLIK) + list(YIL_SUTUN) + [s for s, _, _ in SAYAC]:
        stil(ws[f"{col}4"], sablon["A4"])
    ws.row_dimensions[4].height = 45

    for i, f in enumerate(fonlar):
        r = ILK + i
        degerler = {
            "A": i + 1, "B": f["ticker"], "C": f["name"], "D": f.get("strateji") or f.get("strategy"),
            "E": f["aum_bn_usd"], "U": f["issuer"], "V": tarih_tr(f.get("inception")),
            "W": SIKLIK.get(str(f.get("frequency", "")).lower(), f.get("frequency")),
            "X": f.get("ttm_dps_usd"), "Y": f.get("yield"), "Z": f.get("expense_ratio"),
            "AA": " | ".join(x for x in [f.get("not") or f.get("note"),
                                         "Kaynak: " + f["source_url"] if f.get("source_url") else None] if x),
        }
        for col, v in degerler.items():
            ws[f"{col}{r}"] = v
        # stiller: ABD İlk 20 satır 5'ten
        for col, kaynak in {"A": "A", "B": "B", "C": "C", "D": "C", "E": "E", "U": "U", "V": "U",
                            "W": "U", "X": "W", "Y": "X", "Z": "X", "AA": "AC", "T": "T"}.items():
            stil(ws[f"{col}{r}"], sablon[f"{kaynak}5"])
        ws[f"E{r}"].number_format = "#,##0.0"
        ws[f"X{r}"].number_format = "#,##0.00"
        ws[f"Y{r}"].number_format = "0.0%"
        ws[f"Z{r}"].number_format = "0.00%"
        ws[f"AA{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        for j, col in enumerate(YIL_SUTUN):
            c = ws[f"{col}{r}"]
            c.value = f["years"].get(str(2017 + j))
            stil(c, sablon["F5"])
            if c.value is None:
                c.fill = copy(GRI)
        for sutun, n, ilk_yil in SAYAC:
            c = ws[f"{sutun}{r}"]
            c.value = sayac_formulu(ilk_yil, r)
            stil(c, sablon["P5"])
            c.number_format = "0"
        ws[f"T{r}"] = geriye_kesintisiz(r)
        ws.row_dimensions[r].height = 40

    fon_alt_kisim(ws, sablon, son, [veri.get("selection_note"), veri.get("turkey_note")], FON_GENISLIK)
    return len(fonlar)


def fon_alt_kisim(ws, sablon, son, notlar, genislikler):
    """Fon sayfalarının ortak kısmı: renkler, alt toplam, açıklamalar, sütun genişlikleri."""
    stil = stil_kopyala
    ws.conditional_formatting.add(
        f"F{ILK}:O{son}", FormulaRule(formula=["AND(ISNUMBER(F5),F5=1)"], fill=YESIL))
    ws.conditional_formatting.add(
        f"F{ILK}:O{son}", FormulaRule(formula=["AND(ISNUMBER(F5),F5=0)"], fill=PEMBE))
    sayac_renkleri(ws, ILK, son)

    alt = son + 2
    ws[f"C{alt}"] = "Hiç aksatmadan ödeyen fon sayısı:"
    stil(ws[f"C{alt}"], sablon["C26"])
    for sutun, n, _ in SAYAC:
        ws[f"{sutun}{alt}"] = f"=COUNTIF({sutun}{ILK}:{sutun}{son},{n})"
        stil(ws[f"{sutun}{alt}"], sablon["P26"])
    ws[f"C{alt + 1}"] = ("✓ = o yıl temettü dağıtıldı, ✗ = dağıtılmadı, gri = fon henüz kurulmamıştı "
                         "(sayılmaz). Sayı sütunları: o dönemde kaç yıl temettü dağıtıldığı. "
                         "Mavi yazı = elle girilen veri, siyah = formül.")
    stil(ws[f"C{alt + 1}"], sablon["C27"])

    satir = alt + 3
    for metin in notlar:
        if not metin:
            continue
        ws[f"A{satir}"] = metin
        stil(ws[f"A{satir}"], sablon["A29"])
        ws.merge_cells(f"A{satir}:T{satir}")
        ws.row_dimensions[satir].height = 90
        satir += 1

    for col, w in genislikler.items():
        ws.column_dimensions[col].width = w
    for col in YIL_SUTUN:
        ws.column_dimensions[col].width = 6
    for col in "PQRS":
        ws.column_dimensions[col].width = 11
    ws.freeze_panes = "D5"
    ws.sheet_view.zoomScale = sablon.sheet_view.zoomScale


# ---------------------------------------------------------------- Çin fon sayfası
CIN_FON_BASLIK = {
    "A": "Sıra", "B": "Kod", "C": "Fon", "D": "Ne yapar (strateji)",
    "E": "Fon büyüklüğü\n(milyon USD)",
    "T": "Geriye doğru\nkesintisiz yıl\n(tabloda, max 10)",
    "U": "İhraççı", "V": "Borsa", "W": "Kuruluş", "X": "Ödeme\nsıklığı",
    "Y": "Son 12 ay\ntemettü / pay", "Z": "Para\nbirimi", "AA": "Temettü\nverimi",
    "AB": "Yıllık gider\noranı", "AC": "Lot\n(en az adet)", "AD": "Not / Kaynak",
}
CIN_FON_GENISLIK = {"A": 5, "B": 10, "C": 34, "D": 34, "E": 12, "T": 12, "U": 16, "V": 10,
                    "W": 11, "X": 10, "Y": 11, "Z": 8, "AA": 9, "AB": 10, "AC": 10, "AD": 60}


def cin_fon_sayfasi(wb, veri):
    sablon = wb["ABD İlk 20"]
    ws = wb.create_sheet("Çin Temettü Fonları", index=wb.sheetnames.index("Çin İlk 20") + 1)
    fonlar = veri["funds"]
    son = 4 + len(fonlar)
    stil = stil_kopyala

    ws["A1"] = f"Çin — temettü dağıtan en büyük {len(fonlar)} borsa yatırım fonu (Hong Kong / ABD'de işlem görenler)"
    stil(ws["A1"], sablon["A1"])
    ws["A2"] = (f"Veri tarihi: {veri['as_of']}. Fon = içinde onlarca hisse olan sepet; tek şirkete göre "
                "riski dağıtır. Seçim kriteri ve elenen fonlar tablonun altında.")
    stil(ws["A2"], sablon["A2"])

    for col, metin in CIN_FON_BASLIK.items():
        ws[f"{col}4"] = metin
    for i, col in enumerate(YIL_SUTUN):
        ws[f"{col}4"] = 2017 + i
    for sutun, n, _ in SAYAC:
        ws[f"{sutun}4"] = baslik(n)
    for col in list(CIN_FON_BASLIK) + list(YIL_SUTUN) + [x for x, _, _ in SAYAC]:
        stil(ws[f"{col}4"], sablon["A4"])
    ws.row_dimensions[4].height = 45

    for i, f in enumerate(fonlar):
        r = ILK + i
        hk = str(f.get("exchange", "")).upper().startswith("HK")
        degerler = {
            "A": i + 1, "B": f"{f['code']}.HK" if hk else f["code"], "C": f["name"],
            "D": f.get("strateji") or f.get("strategy"), "E": f.get("aum_mn_usd"),
            "U": f.get("issuer"), "V": "Hong Kong" if hk else f.get("exchange"),
            "W": tarih_tr(f.get("inception")),
            "X": SIKLIK.get(str(f.get("frequency", "")).lower(), f.get("frequency")),
            "Y": f.get("ttm_dps"), "Z": f.get("dps_currency"), "AA": f.get("yield"),
            "AB": f.get("expense_ratio"), "AC": f.get("board_lot"),
            "AD": " | ".join(x for x in [f.get("not") or f.get("note"),
                                         "Kaynak: " + f["source_url"] if f.get("source_url") else None] if x),
        }
        for col, v in degerler.items():
            ws[f"{col}{r}"] = v
        for col, kaynak in {"A": "A", "B": "B", "C": "C", "D": "C", "E": "E", "T": "T", "U": "U",
                            "V": "U", "W": "U", "X": "U", "Y": "W", "Z": "U", "AA": "X", "AB": "X",
                            "AC": "U", "AD": "AC"}.items():
            stil(ws[f"{col}{r}"], sablon[f"{kaynak}5"])
        ws[f"E{r}"].number_format = "#,##0"
        ws[f"Y{r}"].number_format = "#,##0.000"
        ws[f"AA{r}"].number_format = "0.0%"
        ws[f"AB{r}"].number_format = "0.00%"
        ws[f"AC{r}"].number_format = "#,##0"
        ws[f"AD{r}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        for j, col in enumerate(YIL_SUTUN):
            c = ws[f"{col}{r}"]
            c.value = f["years"].get(str(2017 + j))
            stil(c, sablon["F5"])
            if c.value is None:
                c.fill = copy(GRI)
        for sutun, n, ilk_yil in SAYAC:
            c = ws[f"{sutun}{r}"]
            c.value = sayac_formulu(ilk_yil, r)
            stil(c, sablon["P5"])
            c.number_format = "0"
        ws[f"T{r}"] = geriye_kesintisiz(r)
        ws.row_dimensions[r].height = 40

    fon_alt_kisim(ws, sablon, son, [veri.get("selection_note"), veri.get("broad_note")],
                  CIN_FON_GENISLIK)
    return len(fonlar)


def ana_rapor():
    """BIST + ABD + ABD temettü fonları."""
    wb = openpyxl.load_workbook(GIRDI)
    sirket_sayfasi(wb["BIST İlk 20"])
    sirket_sayfasi(wb["ABD İlk 20"])
    fon_sayisi = 0
    if FON_VERISI.exists():
        veri = json.loads(FON_VERISI.read_text(encoding="utf-8"))
        fon_sayisi = fon_sayfasi(wb, veri)
    ozet_sayfasi(wb["Özet"], fon_sayisi)
    batmama_sayfasi(wb["Batmama Riski (3 yıl)"])
    wb.save(CIKTI)
    print("Yazıldı:", CIKTI)


def cin_raporu():
    """Çin için ayrı dosya: Çin İlk 20 + Çin Temettü Fonları + Batmama Riski + Not Tablosu."""
    if not CIN_VERISI.exists():
        return
    veri = json.loads(CIN_VERISI.read_text(encoding="utf-8"))
    wb = openpyxl.load_workbook(GIRDI)  # biçim şablonu olarak
    n = cin_sayfasi(wb, veri)
    if CIN_FON_VERISI.exists():
        cin_fon_sayfasi(wb, json.loads(CIN_FON_VERISI.read_text(encoding="utf-8")))
    bat = wb["Batmama Riski (3 yıl)"]
    batmama_sayfasi(bat)
    cin_batmama_sayfasi(bat, n)
    for ad in ("Özet", "BIST İlk 20", "ABD İlk 20"):
        del wb[ad]
    wb.active = 0
    for ws in wb:
        ws.sheet_view.tabSelected = ws.title == "Çin İlk 20"
    wb.save(CIN_CIKTI)
    print("Yazıldı:", CIN_CIKTI)


def main():
    ana_rapor()
    cin_raporu()


if __name__ == "__main__":
    main()

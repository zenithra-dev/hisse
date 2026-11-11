"""
Temettü Raporu güncelleyici.

Girdi : kaynak/temettu_raporu_orijinal.xlsx  (BIST ve ABD ilk 20 temettü raporu)
        temettu_fonlari.json                 (temettü dağıtan ABD fonları / ETF verisi)
Çıktı : temettu_raporu.xlsx

Yaptığı değişiklikler:
1. "Son 3/5/7/10 yıl kesintisiz" sütunları Evet/Hayır yerine SAYI gösterir:
   10 yılın 9'unda temettü ödediyse 9 yazar. Hepsinde ödediyse yeşil, eksik varsa pembe.
2. Özet, "Batmama Riski" sayfası ve alt toplamlar yeni sayılara göre çalışır.
3. "ABD Temettü Fonları" sayfası eklenir (aynı ✓/✗ ve sayı mantığıyla).

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
CIKTI = KOK / "temettu_raporu.xlsx"

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


def ozet_sayfasi(ws, fon_sayisi):
    etiket = {5: 3, 6: 5, 7: 7, 8: 10}
    sutun = {5: "P", 6: "Q", 7: "R", 8: "S"}
    ws["A4"] = "Hiç aksatmadan temettü ödeyen"
    for r, n in etiket.items():
        s = sutun[r]
        ws[f"A{r}"] = f"Son {n} yılın hepsinde ödeyen"
        ws[f"B{r}"] = f"=COUNTIF('BIST İlk 20'!{s}5:{s}24,{n})"
        ws[f"C{r}"] = f"=COUNTIF('ABD İlk 20'!{s}5:{s}24,{n})"
        ws[f"D{r}"] = f"=B{r}+C{r}"
        if not fon_sayisi:
            continue
        e = ws[f"E{r}"]
        e.value = f"=COUNTIF('ABD Temettü Fonları'!{s}5:{s}{4 + fon_sayisi},{n})"
        for attr in ("font", "border", "alignment", "number_format"):
            setattr(e, attr, copy(getattr(ws[f"C{r}"], attr)))
    ws["A12"] = ("• BIST İlk 20 / ABD İlk 20: her yıl temettü ödendi mi (✓/✗), son 3-5-7-10 yılda "
                 "kaç yıl ödendiği (ör. 10 yılda 9), verim, kredi notu, borçluluk.")
    if fon_sayisi:
        ws["A13"] = ("• ABD Temettü Fonları: temettü dağıtan en büyük 20 ABD borsa yatırım fonu (ETF); "
                     "aynı ✓/✗ ve sayı mantığı, gider oranı ve verim.")
        ws["A14"] = ("• Batmama Riski (3 yıl): iki borsadaki 40 şirket için basit finansal sağlamlık "
                     "göstergesi.  • Not Tablosu: kredi notlarının kategori karşılıkları.")
    ws["A17"] = ("Yıl kolonlarındaki ✓/✗ hücreleri düzeltilirse tüm sayılar ve göstergeler "
                 "otomatik güncellenir.")
    if not fon_sayisi:
        return
    e4 = ws["E4"]
    e4.value = f"ABD temettü fonları\n({fon_sayisi} fon)"
    for attr in ("font", "fill", "border", "alignment"):
        setattr(e4, attr, copy(getattr(ws["C4"], attr)))
    e4.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["D4"] = "Toplam\n(şirket)"
    ws["D4"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[4].height = 30
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["E"].width = 20


def batmama_sayfasi(ws):
    ws["F4"] = "Son 10 yılda\nkaç yıl temettü"
    ws["G4"] = "Son 5 yılda\nkaç yıl temettü"
    for r in range(5, 45):
        ws[f"F{r}"].number_format = "0"
        ws[f"G{r}"].number_format = "0"
        ws[f"J{r}"] = (f'=IF(AND(E{r}="Yatırım yapılabilir",F{r}=10),"Çok güçlü",'
                       f'IF(OR(E{r}="Yatırım yapılabilir",F{r}=10),"Güçlü",'
                       f'IF(G{r}=5,"Orta","Takip edilmeli")))')
    ws["A47"] = ("• Çok güçlü: yatırım yapılabilir kredi notu (BBB- / Baa3 ve üstü) VE son 10 yılın "
                 "10'unda da temettü.")
    ws["A49"] = "• Orta: yalnızca son 5 yılın 5'inde de temettü."


# ---------------------------------------------------------------- fon sayfası
FON_BASLIK = {
    "A": "Sıra", "B": "Kod", "C": "Fon", "D": "Ne yapar (strateji)",
    "E": "Fon büyüklüğü\n(milyar USD)",
    "T": "Geriye doğru\nkesintisiz yıl\n(tabloda, max 10)",
    "U": "İhraççı", "V": "Kuruluş", "W": "Ödeme\nsıklığı",
    "X": "Son 12 ay\ntemettü / pay\n(USD)", "Y": "Temettü\nverimi", "Z": "Yıllık gider\noranı",
    "AA": "Not / Kaynak",
}
FON_GENISLIK = {"A": 5, "B": 8, "C": 30, "D": 26, "E": 12, "T": 12, "U": 14, "V": 11,
                "W": 10, "X": 11, "Y": 9, "Z": 10, "AA": 60}
SIKLIK = {"monthly": "Aylık", "quarterly": "3 ayda bir", "annual": "Yılda bir",
          "semiannual": "6 ayda bir", "semi-annual": "6 ayda bir"}


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
            "E": f["aum_bn_usd"], "U": f["issuer"], "V": f.get("inception"),
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
        ws.row_dimensions[r].height = 30

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

    notlar = [veri.get("selection_note", ""), veri.get("turkey_note", "")]
    satir = alt + 3
    for metin in notlar:
        if not metin:
            continue
        ws[f"A{satir}"] = metin
        stil(ws[f"A{satir}"], sablon["A29"])
        ws.merge_cells(f"A{satir}:T{satir}")
        ws.row_dimensions[satir].height = 90
        satir += 1

    for col, w in FON_GENISLIK.items():
        ws.column_dimensions[col].width = w
    for col in YIL_SUTUN:
        ws.column_dimensions[col].width = 6
    for col in "PQRS":
        ws.column_dimensions[col].width = 11
    ws.freeze_panes = "D5"
    ws.sheet_view.zoomScale = sablon.sheet_view.zoomScale
    return len(fonlar)


def main():
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


if __name__ == "__main__":
    main()

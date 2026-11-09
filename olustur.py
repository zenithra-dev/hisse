"""
Hisse Senedi Portföy & Temettü Takip Tablosu üreticisi.

Çıktı: hisse_takip.xlsx
Dosya Google E-Tablolar'a içe aktarılınca GOOGLEFINANCE formülleri
fiyatları ve dolar kurunu otomatik günceller (BIST ~15 dk gecikmeli).

Kullanım:
    python3 olustur.py                  # hisse_takip.xlsx üretir
    python3 olustur.py --test out.xlsx  # GOOGLEFINANCE yerine sabit sayılar
                                        # (LibreOffice ile formül testi için)
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# ------------------------------------------------------------------ ayarlar
N_HISSE = 50          # portföy sekmesindeki hisse satırı sayısı
N_ISLEM = 500         # alım-satım sekmesindeki işlem satırı sayısı
N_TEMETTU = 300       # temettü geliri satırı sayısı

P_TOPLAM = 4                          # portföy: TOPLAM satırı
P_ILK = 5                             # portföy: ilk veri satırı
P_SON = P_ILK + N_HISSE - 1

I_ILK = 4                             # işlem: ilk veri satırı
I_SON = I_ILK + N_ISLEM - 1

T_ILK = 4                             # temettü: ilk veri satırı
T_SON = T_ILK + N_TEMETTU - 1

# Toplam Adet sekmesi: BIST bloğu portföy ile aynı satırlarda,
# ABD bloğu altında (satır farkı = ABD_OFSET)
ABD_OFSET = N_HISSE + 5

KUR = "'Özet'!$B$5"   # güncel dolar/TL

# Popüler temettü hisseleri (örnek; tavsiye değildir)
ORNEK_BIST = ["TUPRS", "FROTO", "DOAS", "TTRAK", "EREGL", "AYGAZ",
              "TOASO", "KCHOL", "ENJSA", "TCELL", "BIMAS", "ISMEN"]
ORNEK_ABD = [("KO", "NYSE"), ("PEP", "NASDAQ"), ("JNJ", "NYSE"), ("PG", "NYSE"),
             ("ABBV", "NYSE"), ("MO", "NYSE"), ("O", "NYSE"), ("VZ", "NYSE"),
             ("XOM", "NYSE"), ("CVX", "NYSE"), ("MCD", "NYSE"), ("SCHD", "NYSEARCA")]

TEST_MODU = False     # --test ile True olur

# ------------------------------------------------------------------ stiller
KOYU = PatternFill("solid", fgColor="1F3864")
ORTA = PatternFill("solid", fgColor="2E75B6")
YESIL_DOLGU = PatternFill("solid", fgColor="38761D")
GIRIS = PatternFill("solid", fgColor="FFF2CC")       # sarı: kullanıcı yazar
OTO_GIRIS = PatternFill("solid", fgColor="E2EFDA")   # yeşil: otomatik, istenirse elle yazılır
TOPLAM_DOLGU = PatternFill("solid", fgColor="D9E1F2")
BEYAZ_KALIN = Font(bold=True, color="FFFFFF")
KALIN = Font(bold=True)
BASLIK_FONT = Font(bold=True, size=14, color="1F3864")
IPUCU_FONT = Font(italic=True, size=9, color="595959")
INCE = Side(style="thin", color="BFBFBF")
KENAR = Border(left=INCE, right=INCE, top=INCE, bottom=INCE)
ORTALA = Alignment(horizontal="center", vertical="center", wrap_text=True)

BICIM = {
    "tl": '#,##0.00 "₺"',
    "usd": '"$"#,##0.00',
    "adet": '#,##0',
    "yuzde": '0.00%',
    "tarih": 'dd.mm.yyyy',
    "sayi": '#,##0.00',
    "kur": '#,##0.0000',
    "metin": None,
}

YESIL = Font(color="008000", bold=True)
KIRMIZI = Font(color="C00000", bold=True)


def renklendir(ws, aralik, ilk_hucre):
    """Pozitif yeşil, negatif kırmızı (yalnızca sayı olan hücreler)."""
    ws.conditional_formatting.add(
        aralik, FormulaRule(formula=[f"AND(ISNUMBER({ilk_hucre}),{ilk_hucre}>0)"], font=YESIL))
    ws.conditional_formatting.add(
        aralik, FormulaRule(formula=[f"AND(ISNUMBER({ilk_hucre}),{ilk_hucre}<0)"], font=KIRMIZI))


def sutunlar(tanim):
    """[(anahtar, başlık, genişlik, biçim, dolgu)] -> ({anahtar: harf}, tanım)"""
    return {t[0]: get_column_letter(i) for i, t in enumerate(tanim, start=1)}


def baslik_satiri(ws, satir, tanim, dolgu=KOYU):
    for i, t in enumerate(tanim, start=1):
        c = ws.cell(row=satir, column=i, value=t[1])
        c.fill = dolgu
        c.font = BEYAZ_KALIN
        c.alignment = ORTALA
        c.border = KENAR
        ws.column_dimensions[get_column_letter(i)].width = t[2]
    ws.row_dimensions[satir].height = 45


def satir_bicimle(ws, r, tanim, giris_dolgusu=True):
    for i, t in enumerate(tanim, start=1):
        c = ws.cell(row=r, column=i)
        c.border = KENAR
        if BICIM[t[3]]:
            c.number_format = BICIM[t[3]]
        if giris_dolgusu and len(t) > 4 and t[4] is not None:
            c.fill = t[4]


def sayfa_basligi(ws, baslik, ipucu, son_sutun):
    ws["A1"] = baslik
    ws["A1"].font = BASLIK_FONT
    ws["A2"] = ipucu
    ws["A2"].font = IPUCU_FONT
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=min(son_sutun, 14))
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 40


def gf(sembol_ifadesi, ozellik):
    """GOOGLEFINANCE çağrısı (test modunda sabit değer)."""
    if TEST_MODU:
        sabit = {"price": "100", "changepct": "1.5", "high": "105", "low": "95",
                 "high52": "150", "low52": "60", "pe": "8.5"}
        if ozellik == "name":
            return f'"Test "&{sembol_ifadesi}'
        return sabit[ozellik]
    return f'GOOGLEFINANCE({sembol_ifadesi},"{ozellik}")'


def gecmis_kur(tarih_hucre):
    """Verilen tarihteki dolar/TL kapanışı; bulunamazsa güncel kur."""
    if TEST_MODU:
        return "30"
    return (f'IFERROR(INDEX(GOOGLEFINANCE("CURRENCY:USDTRY","price",{tarih_hucre}),2,2),'
            f'{KUR})')


# ------------------------------------------------------------------ Portföy
def portfoy_tanimi(piyasa):
    p = "tl" if piyasa == "BIST" else "usd"
    t = [("kod", "Hisse Kodu", 11, "metin", GIRIS)]
    if piyasa == "ABD":
        t.append(("borsa", "Borsa", 11, "metin", GIRIS))
    t += [
        ("sirket", "Şirket", 26, "metin"),
        ("fiyat", "Güncel Fiyat", 12, p),
        ("degisim", "Günlük Değişim", 10, "yuzde"),
        ("yuksek", "Gün En Yüksek", 11, p),
        ("dusuk", "Gün En Düşük", 11, p),
        ("adet", "Eldeki Adet", 10, "adet"),
        ("ortmal", "Ortalama Maliyet", 12, p),
        ("topmal", "Toplam Maliyet", 14, p),
        ("deger", "Güncel Değer", 14, p),
        ("kz", "Kâr / Zarar", 14, p),
        ("kzp", "Kâr / Zarar %", 10, "yuzde"),
        ("pay", "Portföy Payı", 9, "yuzde"),
        ("temettu", "Hisse Başı Yıllık Temettü", 12, p, GIRIS),
        ("verim", "Temettü Verimi", 10, "yuzde"),
        ("temgelir", "Yıllık Temettü Geliri", 13, p),
        ("mverim", "Maliyete Göre Temettü Verimi", 12, "yuzde"),
    ]
    if piyasa == "ABD":
        t += [
            ("fiyat_tl", "Güncel Fiyat (₺)", 13, "tl"),
            ("topmal_tl", "Toplam Maliyet (₺)\nalış günü kuruyla", 16, "tl"),
            ("deger_tl", "Güncel Değer (₺)", 16, "tl"),
            ("kz_tl", "Kâr / Zarar (₺)", 15, "tl"),
            ("kzp_tl", "Kâr / Zarar % (₺)", 10, "yuzde"),
            ("temgelir_tl", "Yıllık Temettü Geliri (₺)", 14, "tl"),
        ]
    t += [
        ("y52h", "52 Hafta En Yüksek", 11, p),
        ("y52l", "52 Hafta En Düşük", 11, p),
        ("fk", "F/K", 7, "sayi"),
    ]
    return t


def portfoy_sayfasi(wb, ad, islem_sayfa, piyasa):
    ws = wb.create_sheet(ad)
    tanim = portfoy_tanimi(piyasa)
    C = sutunlar(tanim)
    ofset = 0 if piyasa == "BIST" else ABD_OFSET

    if piyasa == "BIST":
        baslik = "BORSA İSTANBUL (₺) — PORTFÖY & TEMETTÜ"
        kod_ipucu = "Hisse kodunu IST: olmadan yazın (ör. THYAO)."
    else:
        baslik = "ABD BORSALARI ($ + ₺) — PORTFÖY & TEMETTÜ"
        kod_ipucu = ("Hisse kodunu yazın (ör. KO) ve Borsa sütunundan borsasını seçin "
                     "(NYSE / NASDAQ / NYSEARCA). TL sütunları dolar kuruyla otomatik hesaplanır.")
    sayfa_basligi(
        ws, baslik,
        f"Sarı hücrelere siz yazarsınız, gerisi otomatik. {kod_ipucu} "
        f"Adet ve maliyet '{islem_sayfa}' sekmesindeki kayıtlardan hesaplanır. "
        "Yıllık temettü tutarını (hisse başı) kendiniz girin. Listedeki hisseler örnektir, tavsiye değildir.",
        len(tanim))
    baslik_satiri(ws, 3, tanim)

    # TOPLAM satırı
    t = P_TOPLAM
    ws[f"A{t}"] = "TOPLAM"
    for k in ["topmal", "deger", "temgelir"] + (
            ["topmal_tl", "deger_tl", "temgelir_tl"] if piyasa == "ABD" else []):
        ws[f"{C[k]}{t}"] = f"=SUM({C[k]}{P_ILK}:{C[k]}{P_SON})"
    ws[f"{C['kz']}{t}"] = f"={C['deger']}{t}-{C['topmal']}{t}"
    ws[f"{C['kzp']}{t}"] = f'=IF({C["topmal"]}{t}=0,"",{C["kz"]}{t}/{C["topmal"]}{t})'
    ws[f"{C['pay']}{t}"] = f'=IF({C["deger"]}{t}=0,"",1)'
    ws[f"{C['verim']}{t}"] = f'=IF({C["deger"]}{t}=0,"",{C["temgelir"]}{t}/{C["deger"]}{t})'
    ws[f"{C['mverim']}{t}"] = f'=IF({C["topmal"]}{t}=0,"",{C["temgelir"]}{t}/{C["topmal"]}{t})'
    if piyasa == "ABD":
        ws[f"{C['kz_tl']}{t}"] = f"={C['deger_tl']}{t}-{C['topmal_tl']}{t}"
        ws[f"{C['kzp_tl']}{t}"] = (f'=IF({C["topmal_tl"]}{t}=0,"",'
                                   f'{C["kz_tl"]}{t}/{C["topmal_tl"]}{t})')
        ws[f"{C['fiyat_tl']}{t}"] = f'="1$ = "&TEXT({KUR},"0.0000")&" ₺"'
    satir_bicimle(ws, t, tanim, giris_dolgusu=False)
    for i in range(1, len(tanim) + 1):
        ws.cell(row=t, column=i).fill = TOPLAM_DOLGU
        ws.cell(row=t, column=i).font = KALIN

    ta = "'Toplam Adet'"
    for r in range(P_ILK, P_SON + 1):
        h = {k: f"{v}{r}" for k, v in C.items()}
        a = h["kod"]
        bos = f'{a}=""'
        if piyasa == "BIST":
            sembol = f'"IST:"&{a}'
        else:
            sembol = f'IF({h["borsa"]}="",{a},{h["borsa"]}&":"&{a})'
        tr = r + ofset   # Toplam Adet sekmesindeki satır
        f = {
            "sirket": f'=IF({bos},"",IFERROR({gf(sembol, "name")},"⚠ Kod bulunamadı"))',
            "fiyat": f'=IF({bos},"",IFERROR({gf(sembol, "price")},""))',
            "degisim": f'=IF({bos},"",IFERROR({gf(sembol, "changepct")}/100,""))',
            "yuksek": f'=IF({bos},"",IFERROR({gf(sembol, "high")},""))',
            "dusuk": f'=IF({bos},"",IFERROR({gf(sembol, "low")},""))',
            "adet": f'=IF({bos},"",N({ta}!H{tr}))',
            "ortmal": f'=IF(OR({bos},N({h["adet"]})=0),"",{ta}!E{tr})',
            "topmal": f'=IF({h["ortmal"]}="","",{h["adet"]}*{h["ortmal"]})',
            "deger": f'=IF(OR({bos},{h["fiyat"]}="",N({h["adet"]})=0),"",{h["adet"]}*{h["fiyat"]})',
            "kz": f'=IF(OR({h["topmal"]}="",{h["deger"]}=""),"",{h["deger"]}-{h["topmal"]})',
            "kzp": f'=IF({h["kz"]}="","",IFERROR({h["kz"]}/{h["topmal"]},""))',
            "pay": (f'=IF(OR({h["deger"]}="",N(${C["deger"]}${t})=0),"",'
                    f'{h["deger"]}/${C["deger"]}${t})'),
            "verim": (f'=IF(OR({h["temettu"]}="",{h["fiyat"]}=""),"",'
                      f'IFERROR({h["temettu"]}/{h["fiyat"]},""))'),
            "temgelir": (f'=IF(OR({h["temettu"]}="",N({h["adet"]})=0),"",'
                         f'{h["temettu"]}*{h["adet"]})'),
            "mverim": (f'=IF(OR({h["temettu"]}="",{h["ortmal"]}=""),"",'
                       f'IFERROR({h["temettu"]}/{h["ortmal"]},""))'),
            "y52h": f'=IF({bos},"",IFERROR({gf(sembol, "high52")},""))',
            "y52l": f'=IF({bos},"",IFERROR({gf(sembol, "low52")},""))',
            "fk": f'=IF({bos},"",IFERROR({gf(sembol, "pe")},""))',
        }
        if piyasa == "ABD":
            f.update({
                "fiyat_tl": f'=IF(OR({h["fiyat"]}="",N({KUR})=0),"",{h["fiyat"]}*{KUR})',
                "topmal_tl": f'=IF({h["ortmal"]}="","",{h["adet"]}*{ta}!K{tr})',
                "deger_tl": f'=IF(OR({h["deger"]}="",N({KUR})=0),"",{h["deger"]}*{KUR})',
                "kz_tl": (f'=IF(OR({h["topmal_tl"]}="",{h["deger_tl"]}=""),"",'
                          f'{h["deger_tl"]}-{h["topmal_tl"]})'),
                "kzp_tl": f'=IF({h["kz_tl"]}="","",IFERROR({h["kz_tl"]}/{h["topmal_tl"]},""))',
                "temgelir_tl": (f'=IF(OR({h["temgelir"]}="",N({KUR})=0),"",'
                                f'{h["temgelir"]}*{KUR})'),
            })
        for k, formul in f.items():
            ws[h[k]] = formul
        satir_bicimle(ws, r, tanim)
        ws[a].font = KALIN

    # örnek hisseler
    for i, ornek in enumerate(ORNEK_BIST if piyasa == "BIST" else ORNEK_ABD):
        r = P_ILK + i
        if piyasa == "BIST":
            ws[f"{C['kod']}{r}"] = ornek
        else:
            ws[f"{C['kod']}{r}"] = ornek[0]
            ws[f"{C['borsa']}{r}"] = ornek[1]

    if piyasa == "ABD":
        dv = DataValidation(type="list", formula1='"NYSE,NASDAQ,NYSEARCA,NYSEAMERICAN"',
                            allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{C['borsa']}{P_ILK}:{C['borsa']}{P_SON}")

    for k in ["degisim"]:
        renklendir(ws, f"{C[k]}{P_TOPLAM}:{C[k]}{P_SON}", f"{C[k]}{P_TOPLAM}")
    renklendir(ws, f"{C['kz']}{P_TOPLAM}:{C['kzp']}{P_SON}", f"{C['kz']}{P_TOPLAM}")
    if piyasa == "ABD":
        renklendir(ws, f"{C['kz_tl']}{P_TOPLAM}:{C['kzp_tl']}{P_SON}", f"{C['kz_tl']}{P_TOPLAM}")
    ws.freeze_panes = f"{get_column_letter(2 if piyasa == 'BIST' else 3)}{P_ILK}"
    return C


# ------------------------------------------------------------------ Alım-Satım
def islem_tanimi(piyasa):
    p = "tl" if piyasa == "BIST" else "usd"
    t = [
        ("tarih", "Tarih", 12, "tarih", GIRIS),
        ("kod", "Hisse Kodu", 11, "metin", GIRIS),
        ("islem", "İşlem", 9, "metin", GIRIS),
        ("adet", "Adet", 9, "adet", GIRIS),
        ("fiyat", "Fiyat", 11, p, GIRIS),
        ("komisyon", "Komisyon", 10, p, GIRIS),
        ("tutar", "Toplam Tutar", 14, p),
        ("guncel", "Güncel Fiyat", 12, p),
        ("deger", "Güncel Değer", 14, p),
        ("kz", "Kâr / Zarar", 14, p),
        ("kzp", "Kâr / Zarar %", 10, "yuzde"),
    ]
    if piyasa == "ABD":
        t += [
            ("kur", "İşlem Günü Dolar Kuru\n(otomatik)", 12, "kur", OTO_GIRIS),
            ("tutar_tl", "Toplam Tutar (₺)", 15, "tl"),
            ("deger_tl", "Güncel Değer (₺)", 15, "tl"),
            ("kz_tl", "Kâr / Zarar (₺)", 14, "tl"),
            ("kzp_tl", "Kâr / Zarar % (₺)", 10, "yuzde"),
        ]
    t.append(("uyari", "Uyarı", 36, "metin"))
    return t


def islem_sayfasi(wb, ad, portfoy_sayfa, piyasa, PC):
    ws = wb.create_sheet(ad)
    tanim = islem_tanimi(piyasa)
    C = sutunlar(tanim)
    ek = ("Dolar kuru işlem tarihine göre otomatik gelir; isterseniz yeşil hücreye kendi kurunuzu yazabilirsiniz. "
          if piyasa == "ABD" else "")
    sayfa_basligi(
        ws, f"{'BORSA İSTANBUL' if piyasa == 'BIST' else 'ABD'} — ALIM / SATIM KAYITLARI",
        "Her alım veya satımı bir satıra yazın (sarı sütunlar). İşlem sütunundan ALIŞ / SATIŞ seçin. "
        "Komisyon boş bırakılabilir. Kâr/zarar yalnızca ALIŞ satırlarında, güncel fiyata göre gösterilir; "
        f"satışlardan gerçekleşen kâr 'Toplam Adet' sekmesindedir. {ek}",
        len(tanim))
    baslik_satiri(ws, 3, tanim)

    pf = f"'{portfoy_sayfa}'"
    pkod = f"{pf}!${PC['kod']}${P_ILK}:${PC['kod']}${P_SON}"
    # VLOOKUP için kod..fiyat aralığı ve fiyatın sütun sırası
    fiyat_sira = list(PC).index("fiyat") + 1
    paralik = f"{pf}!${PC['kod']}${P_ILK}:${PC['fiyat']}${P_SON}"

    dv = DataValidation(type="list", formula1='"ALIŞ,SATIŞ"', allow_blank=True,
                        showErrorMessage=True, errorTitle="Geçersiz",
                        error="ALIŞ veya SATIŞ seçin.")
    ws.add_data_validation(dv)
    dv.add(f"{C['islem']}{I_ILK}:{C['islem']}{I_SON}")

    for r in range(I_ILK, I_SON + 1):
        h = {k: f"{v}{r}" for k, v in C.items()}
        f = {
            "tutar": (f'=IF(OR({h["kod"]}="",{h["adet"]}="",{h["fiyat"]}=""),"",'
                      f'IF({h["islem"]}="SATIŞ",{h["adet"]}*{h["fiyat"]}-N({h["komisyon"]}),'
                      f'{h["adet"]}*{h["fiyat"]}+N({h["komisyon"]})))'),
            "guncel": (f'=IF({h["kod"]}="","",IFERROR(VLOOKUP({h["kod"]},{paralik},'
                       f'{fiyat_sira},FALSE),""))'),
            "deger": (f'=IF(OR({h["islem"]}<>"ALIŞ",{h["tutar"]}="",{h["guncel"]}=""),"",'
                      f'{h["adet"]}*{h["guncel"]})'),
            "kz": f'=IF({h["deger"]}="","",{h["deger"]}-{h["tutar"]})',
            "kzp": f'=IF({h["kz"]}="","",IFERROR({h["kz"]}/{h["tutar"]},""))',
            "uyari": (f'=IF({h["kod"]}="","",IF(COUNTIF({pkod},{h["kod"]})=0,'
                      f'"⚠ Bu hisseyi \'{portfoy_sayfa}\' sekmesine ekleyin",'
                      f'IF({h["islem"]}="","⚠ İşlem türü seçin (ALIŞ/SATIŞ)",'
                      f'IF({h["tarih"]}="","⚠ Tarih yazın",""))))'),
        }
        if piyasa == "ABD":
            f.update({
                "kur": f'=IF(OR({h["kod"]}="",{h["tarih"]}=""),"",{gecmis_kur(h["tarih"])})',
                "tutar_tl": (f'=IF(OR({h["tutar"]}="",{h["kur"]}=""),"",'
                             f'{h["tutar"]}*{h["kur"]})'),
                "deger_tl": (f'=IF(OR({h["deger"]}="",N({KUR})=0),"",'
                             f'{h["deger"]}*{KUR})'),
                "kz_tl": (f'=IF(OR({h["deger_tl"]}="",{h["tutar_tl"]}=""),"",'
                          f'{h["deger_tl"]}-{h["tutar_tl"]})'),
                "kzp_tl": (f'=IF({h["kz_tl"]}="","",IFERROR({h["kz_tl"]}/{h["tutar_tl"]},""))'),
            })
        for k, formul in f.items():
            ws[h[k]] = formul
        satir_bicimle(ws, r, tanim)
        ws[h["uyari"]].font = Font(color="C00000", size=9)
        ws[h["islem"]].alignment = Alignment(horizontal="center")

    renklendir(ws, f"{C['kz']}{I_ILK}:{C['kzp']}{I_SON}", f"{C['kz']}{I_ILK}")
    if piyasa == "ABD":
        renklendir(ws, f"{C['kz_tl']}{I_ILK}:{C['kzp_tl']}{I_SON}", f"{C['kz_tl']}{I_ILK}")
    ws.freeze_panes = f"C{I_ILK}"
    return C


# ------------------------------------------------------------------ Toplam Adet
def toplam_adet_tanimi(piyasa):
    p = "tl" if piyasa == "BIST" else "usd"
    t = [
        ("kod", "Hisse Kodu", 11, "metin"),           # A
        ("alim_sayisi", "Alım Sayısı", 9, "adet"),     # B
        ("alinan", "Toplam Alınan Adet", 11, "adet"),  # C
        ("alis", "Toplam Alış Tutarı", 15, p),         # D
        ("ort", "Ortalama Alış Fiyatı", 13, p),        # E
        ("satilan", "Toplam Satılan Adet", 11, "adet"),  # F
        ("satis", "Toplam Satış Tutarı", 15, p),       # G
        ("eldeki", "Eldeki Adet", 11, "adet"),         # H
        ("gkz", "Satıştan Gerçekleşen Kâr / Zarar", 16, p),  # I
    ]
    if piyasa == "ABD":
        t += [
            ("alis_tl", "Toplam Alış Tutarı (₺)", 16, "tl"),        # J
            ("ort_tl", "Ortalama Alış Fiyatı (₺)", 14, "tl"),       # K
            ("satis_tl", "Toplam Satış Tutarı (₺)", 16, "tl"),      # L
            ("gkz_tl", "Satıştan Gerçekleşen Kâr / Zarar (₺)", 17, "tl"),  # M
        ]
    return t


def toplam_adet_sayfasi(wb, IC_BIST, IC_ABD):
    ws = wb.create_sheet("Toplam Adet")
    sayfa_basligi(
        ws, "TOPLAM ADET — ALIM / SATIM ÖZETİ",
        "Bu sekme tamamen otomatiktir; bir şey yazmanıza gerek yok. Hisse listesi Portföy sekmelerinden gelir. "
        "Ortalama alış fiyatı = tüm alışların toplam tutarı (komisyon dahil) ÷ toplam alınan adet. "
        "ABD için TL tutarlar işlem günündeki dolar kuruyla hesaplanır.",
        13)

    def blok(baslik_satir, portfoy_sayfa, islem_sayfa, piyasa, IC):
        ofset = 0 if piyasa == "BIST" else ABD_OFSET
        tanim = toplam_adet_tanimi(piyasa)
        C = sutunlar(tanim)
        ws.cell(row=baslik_satir - 1, column=1, value=(
            "BORSA İSTANBUL (₺)" if piyasa == "BIST" else "ABD BORSALARI ($ + ₺)")).font = BASLIK_FONT
        baslik_satiri(ws, baslik_satir, tanim, dolgu=KOYU if piyasa == "BIST" else ORTA)
        top = baslik_satir + 1
        ilk, son = P_ILK + ofset, P_SON + ofset
        ws[f"A{top}"] = "TOPLAM"
        for k in ["alim_sayisi", "alinan", "alis", "satilan", "satis", "eldeki", "gkz"] + (
                ["alis_tl", "satis_tl", "gkz_tl"] if piyasa == "ABD" else []):
            ws[f"{C[k]}{top}"] = f"=SUM({C[k]}{ilk}:{C[k]}{son})"
        satir_bicimle(ws, top, tanim)
        for i in range(1, len(tanim) + 1):
            ws.cell(row=top, column=i).fill = TOPLAM_DOLGU
            ws.cell(row=top, column=i).font = KALIN

        pf = f"'{portfoy_sayfa}'"
        isl = f"'{islem_sayfa}'"

        def ar(k):
            return f"{isl}!${IC[k]}${I_ILK}:${IC[k]}${I_SON}"
        hisse, tur = ar("kod"), ar("islem")
        for r in range(ilk, son + 1):
            h = {k: f"{v}{r}" for k, v in C.items()}
            bos = f'{h["kod"]}=""'

            def topla(kaynak, islem):
                return f'=IF({bos},"",SUMIFS({ar(kaynak)},{hisse},{h["kod"]},{tur},"{islem}"))'
            f = {
                "kod": f'=IF({pf}!A{r - ofset}="","",{pf}!A{r - ofset})',
                "alim_sayisi": f'=IF({bos},"",COUNTIFS({hisse},{h["kod"]},{tur},"ALIŞ"))',
                "alinan": topla("adet", "ALIŞ"),
                "alis": topla("tutar", "ALIŞ"),
                "ort": f'=IF(OR({bos},N({h["alinan"]})=0),"",{h["alis"]}/{h["alinan"]})',
                "satilan": topla("adet", "SATIŞ"),
                "satis": topla("tutar", "SATIŞ"),
                "eldeki": f'=IF({bos},"",{h["alinan"]}-{h["satilan"]})',
                "gkz": (f'=IF(OR({bos},N({h["satilan"]})=0,{h["ort"]}=""),"",'
                        f'{h["satis"]}-{h["satilan"]}*{h["ort"]})'),
            }
            if piyasa == "ABD":
                f.update({
                    "alis_tl": topla("tutar_tl", "ALIŞ"),
                    "ort_tl": f'=IF(OR({bos},N({h["alinan"]})=0),"",{h["alis_tl"]}/{h["alinan"]})',
                    "satis_tl": topla("tutar_tl", "SATIŞ"),
                    "gkz_tl": (f'=IF(OR({bos},N({h["satilan"]})=0,{h["ort_tl"]}=""),"",'
                               f'{h["satis_tl"]}-{h["satilan"]}*{h["ort_tl"]})'),
                })
            for k, formul in f.items():
                ws[h[k]] = formul
            satir_bicimle(ws, r, tanim)
            ws[h["kod"]].font = KALIN
        renklendir(ws, f"{C['gkz']}{top}:{C['gkz']}{son}", f"{C['gkz']}{top}")
        if piyasa == "ABD":
            renklendir(ws, f"{C['gkz_tl']}{top}:{C['gkz_tl']}{son}", f"{C['gkz_tl']}{top}")
        return C

    CB = blok(3, "BIST Portföy", "BIST Alım-Satım", "BIST", IC_BIST)
    CA = blok(3 + ABD_OFSET, "ABD Portföy", "ABD Alım-Satım", "ABD", IC_ABD)
    # Portföy formülleri sabit harfler kullanıyor (E, H, K); doğrula
    assert CB["ort"] == "E" and CB["eldeki"] == "H" and CA["ort_tl"] == "K"
    ws.freeze_panes = "B4"
    return CB, CA


# ------------------------------------------------------------------ Temettü Gelirleri
def temettu_sayfasi(wb):
    ws = wb.create_sheet("Temettü Gelirleri")
    tanim = [
        ("tarih", "Ödeme Tarihi", 12, "tarih", GIRIS),
        ("piyasa", "Piyasa", 9, "metin", GIRIS),
        ("kod", "Hisse Kodu", 11, "metin", GIRIS),
        ("adet", "Adet", 10, "adet", GIRIS),
        ("hbt", "Hisse Başı Net Temettü", 13, "kur", GIRIS),
        ("toplam", "Toplam Alınan", 14, "sayi"),
        ("pb", "Para Birimi", 8, "metin"),
        ("kur", "Dolar Kuru\n(ABD için, otomatik)", 12, "kur", OTO_GIRIS),
        ("tl", "TL Karşılığı", 14, "tl"),
    ]
    C = sutunlar(tanim)
    sayfa_basligi(
        ws, "ALINAN TEMETTÜLER",
        "Hesabınıza yatan her temettüyü bir satıra yazın (sarı sütunlar). Piyasa sütunundan BIST / ABD seçin. "
        "ABD temettülerinin TL karşılığı ödeme günündeki dolar kuruyla hesaplanır. Toplamlar 'Özet' sekmesine gider.",
        len(tanim))
    baslik_satiri(ws, 3, tanim)
    dv = DataValidation(type="list", formula1='"BIST,ABD"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"{C['piyasa']}{T_ILK}:{C['piyasa']}{T_SON}")
    for r in range(T_ILK, T_SON + 1):
        h = {k: f"{v}{r}" for k, v in C.items()}
        ws[h["toplam"]] = f'=IF(OR({h["adet"]}="",{h["hbt"]}=""),"",{h["adet"]}*{h["hbt"]})'
        ws[h["pb"]] = f'=IF({h["piyasa"]}="BIST","₺",IF({h["piyasa"]}="ABD","$",""))'
        ws[h["kur"]] = (f'=IF({h["piyasa"]}="BIST",1,IF(AND({h["piyasa"]}="ABD",{h["tarih"]}<>""),'
                        f'{gecmis_kur(h["tarih"])},IF({h["piyasa"]}="ABD",{KUR},"")))')
        ws[h["tl"]] = f'=IF(OR({h["toplam"]}="",{h["kur"]}=""),"",{h["toplam"]}*{h["kur"]})'
        satir_bicimle(ws, r, tanim)
        ws[h["piyasa"]].alignment = Alignment(horizontal="center")
        ws[h["pb"]].alignment = Alignment(horizontal="center")
    ws.freeze_panes = f"A{T_ILK}"
    return C


# ------------------------------------------------------------------ Özet
def ozet_sayfasi(ws, PB, PA, TB, TA, TC):
    ws["A1"] = "HİSSE SENEDİ PORTFÖY & TEMETTÜ TAKİBİ — ÖZET"
    ws["A1"].font = Font(bold=True, size=16, color="1F3864")
    ws["A2"] = ("Fiyatlar ve dolar kuru Google Finance'ten otomatik gelir (BIST yaklaşık 15 dk gecikmeli). "
                "Bu sekmede bir şey yazmanıza gerek yok. Kullanım için 'Nasıl Kullanılır' sekmesine bakın.")
    ws["A2"].font = IPUCU_FONT
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:E2")
    ws.row_dimensions[2].height = 30

    ws["A4"] = "PİYASA"
    ws["A4"].font = KALIN
    gostergeler = [
        ("Dolar / TL", "CURRENCY:USDTRY", '#,##0.0000 "₺"', "40"),
        ("BIST 100", "INDEXIST:XU100", BICIM["sayi"], "10000"),
        ("S&P 500", "INDEXSP:.INX", BICIM["sayi"], "6000"),
    ]
    for i, (ad, sembol, bicim, test) in enumerate(gostergeler):
        r = 5 + i
        ws[f"A{r}"] = ad
        ws[f"B{r}"] = f"={test}" if TEST_MODU else f'=IFERROR(GOOGLEFINANCE("{sembol}"),"-")'
        ws[f"B{r}"].number_format = bicim
        ws[f"B{r}"].font = KALIN
        for c in "AB":
            ws[f"{c}{r}"].border = KENAR
    assert KUR == "'Özet'!$B$5"

    bas = 10
    tanim = [("ad", "", 36), ("b", "BIST (₺)", 18), ("a", "ABD ($)", 18),
             ("atl", "ABD (₺)", 18), ("top", "TOPLAM (₺)", 18)]
    baslik_satiri(ws, bas, tanim)
    bp, ap, ta, tm = "'BIST Portföy'", "'ABD Portföy'", "'Toplam Adet'", "'Temettü Gelirleri'"
    t = P_TOPLAM
    at = 4 + ABD_OFSET
    tem = lambda col, pz: (f'=SUMIFS({tm}!${TC[col]}${T_ILK}:${TC[col]}${T_SON},'
                           f'{tm}!${TC["piyasa"]}${T_ILK}:${TC["piyasa"]}${T_SON},"{pz}")')
    # (ad, BIST ₺, ABD $, ABD ₺)
    satirlar = [
        ("Toplam Maliyet", f"={bp}!{PB['topmal']}{t}", f"={ap}!{PA['topmal']}{t}",
         f"={ap}!{PA['topmal_tl']}{t}"),
        ("Güncel Değer", f"={bp}!{PB['deger']}{t}", f"={ap}!{PA['deger']}{t}",
         f"={ap}!{PA['deger_tl']}{t}"),
        ("Kâr / Zarar (eldeki hisseler)", f"={bp}!{PB['kz']}{t}", f"={ap}!{PA['kz']}{t}",
         f"={ap}!{PA['kz_tl']}{t}"),
        ("Kâr / Zarar %", "%", "%", "%"),
        ("Satışlardan Gerçekleşen Kâr / Zarar", f"={ta}!{TB['gkz']}4",
         f"={ta}!{TA['gkz']}{at}", f"={ta}!{TA['gkz_tl']}{at}"),
        ("Tahmini Yıllık Temettü Geliri", f"={bp}!{PB['temgelir']}{t}",
         f"={ap}!{PA['temgelir']}{t}", f"={ap}!{PA['temgelir_tl']}{t}"),
        ("Portföy Temettü Verimi", "v", "v", "v"),
        ("Bugüne Kadar Alınan Temettü", tem("toplam", "BIST"), tem("toplam", "ABD"),
         tem("tl", "ABD")),
    ]
    for i, (ad, b, a, atl) in enumerate(satirlar):
        r = bas + 1 + i
        ws[f"A{r}"] = ad
        ws[f"A{r}"].font = KALIN
        if b == "%":       # K/Z ÷ maliyet
            m, kz = bas + 1, bas + 3
            for c in "BCDE":
                ws[f"{c}{r}"] = f'=IF(N({c}{m})=0,"",{c}{kz}/{c}{m})'
                ws[f"{c}{r}"].number_format = BICIM["yuzde"]
        elif b == "v":     # yıllık temettü ÷ güncel değer
            d, tg = bas + 2, bas + 6
            for c in "BCDE":
                ws[f"{c}{r}"] = f'=IF(N({c}{d})=0,"",{c}{tg}/{c}{d})'
                ws[f"{c}{r}"].number_format = BICIM["yuzde"]
        else:
            ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = b, a, atl
            ws[f"E{r}"] = f"=N(B{r})+N(D{r})"
            ws[f"B{r}"].number_format = BICIM["tl"]
            ws[f"C{r}"].number_format = BICIM["usd"]
            ws[f"D{r}"].number_format = BICIM["tl"]
            ws[f"E{r}"].number_format = BICIM["tl"]
        for c in "ABCDE":
            ws[f"{c}{r}"].border = KENAR
        ws[f"E{r}"].fill = TOPLAM_DOLGU
        ws[f"E{r}"].font = KALIN
    son = bas + len(satirlar)
    renklendir(ws, f"B{bas + 3}:E{bas + 5}", f"B{bas + 3}")

    notlar = [
        "ABD (₺) sütunu: maliyet, alım günündeki dolar kuruyla; güncel değer, bugünkü kurla hesaplanır.",
        "Böylece TL bazında gerçek kâr/zarar (kur farkı dahil) görünür.",
        "Bu tablo bilgi amaçlıdır, yatırım tavsiyesi değildir.",
    ]
    for i, n in enumerate(notlar):
        ws[f"A{son + 2 + i}"] = n
        ws[f"A{son + 2 + i}"].font = IPUCU_FONT


# ------------------------------------------------------------------ Kılavuz
def kilavuz_sayfasi(wb):
    ws = wb.create_sheet("Nasıl Kullanılır")
    ws.column_dimensions["A"].width = 115
    satirlar = [
        ("NASIL KULLANILIR?", "baslik"),
        ("", None),
        ("1) GOOGLE E-TABLOLAR'A YÜKLEME (bir kere yapılır)", "alt"),
        ("• Bilgisayardan sheets.google.com adresine girin ve 'Boş e-tablo' açın.", None),
        ("• Menüden Dosya → İçe aktar → Yükle → bu dosyayı (hisse_takip.xlsx) seçin.", None),
        ("• 'E-tabloyu değiştir' seçeneğini seçip 'Verileri içe aktar' deyin.", None),
        ("• ÖNEMLİ: Dosyayı Drive'a yükleyip doğrudan açarsanız fiyatlar gelmeyebilir. Öyle açtıysanız", None),
        ("  Dosya → 'Google E-Tablolar olarak kaydet' deyin ve yeni açılan kopyayı kullanın.", None),
        ("• Birkaç saniye içinde örnek hisselerin fiyatları görünür. Sonra telefondan 'Google E-Tablolar' uygulamasıyla da açabilirsiniz.", None),
        ("", None),
        ("2) HİSSE EKLEME / ÇIKARMA", "alt"),
        ("• 'BIST Portföy' veya 'ABD Portföy' sekmesinde A sütunundaki sarı hücreye hisse kodunu yazın.", None),
        ("• BIST için sadece kodu yazın: THYAO, EREGL, TUPRS ... (başına IST: eklemeyin).", None),
        ("• ABD için kodu A sütununa (KO, AAPL ...), borsasını B sütununa (NYSE / NASDAQ / NYSEARCA) yazın.", None),
        ("  Borsayı bilmiyorsanız boş bırakın; 'Kod bulunamadı' çıkarsa borsayı seçin.", None),
        ("• Listede popüler temettü hisseleri örnek olarak hazır gelir. İstemediklerinizi silebilir, kendi hisselerinizi ekleyebilirsiniz.", None),
        ("", None),
        ("3) ALIM / SATIM KAYDI", "alt"),
        ("• 'BIST Alım-Satım' veya 'ABD Alım-Satım' sekmesinde her işlemi bir satıra yazın:", None),
        ("  Tarih, Hisse Kodu, İşlem (ALIŞ / SATIŞ listeden seçilir), Adet, Fiyat, Komisyon (isteğe bağlı).", None),
        ("• Hisse kodu Portföy sekmesindekiyle aynı yazılmalı. Değilse 'Uyarı' sütununda kırmızı yazı çıkar.", None),
        ("• ABD işlemlerinde o günün dolar kuru otomatik gelir (yeşil hücre). Kendi kurunuzu yazmak isterseniz üstüne yazabilirsiniz.", None),
        ("• Eldeki adet, ortalama maliyet ve kâr/zarar kendiliğinden hesaplanır.", None),
        ("• 'Toplam Adet' sekmesinde her hisse için toplam alınan / satılan / eldeki adet ve ortalama alış fiyatı görünür.", None),
        ("", None),
        ("4) TEMETTÜ", "alt"),
        ("• Portföy sekmesindeki 'Hisse Başı Yıllık Temettü' sütununa şirketin bir yılda dağıttığı", None),
        ("  hisse başı temettüyü yazın (KAP, şirketin sitesi veya aracı kurum uygulamasından öğrenilir).", None),
        ("  Tablo temettü verimini, maliyete göre verimi ve tahmini yıllık temettü gelirinizi hesaplar.", None),
        ("• Hesabınıza temettü yattığında 'Temettü Gelirleri' sekmesine bir satır ekleyin.", None),
        ("", None),
        ("5) BİLMENİZ GEREKENLER", "alt"),
        ("• Fiyatlar Google Finance'ten gelir. BIST fiyatları yaklaşık 15 dakika gecikmelidir; borsa açıkken birkaç dakikada bir yenilenir.", None),
        ("• Anlık alış/satış kademe fiyatları (derinlik) ücretsiz kaynaklarda yoktur; 'Güncel Fiyat' son işlem fiyatıdır.", None),
        ("• Temettü tutarı Google Finance'ten otomatik gelmez; bu yüzden elle girilir.", None),
        ("• Sarı hücreler sizin yazacağınız yerlerdir. Diğer hücrelerde formül vardır, silmeyin.", None),
        ("• Bu tablo bilgi amaçlıdır, yatırım tavsiyesi değildir.", None),
    ]
    for i, (metin, tip) in enumerate(satirlar, start=1):
        c = ws.cell(row=i, column=1, value=metin)
        if tip == "baslik":
            c.font = BASLIK_FONT
        elif tip == "alt":
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = ORTA
        c.alignment = Alignment(wrap_text=True, vertical="top")


# ------------------------------------------------------------------ ana
def olustur(cikti):
    wb = Workbook()
    ozet = wb.active
    ozet.title = "Özet"
    PB = portfoy_sayfasi(wb, "BIST Portföy", "BIST Alım-Satım", "BIST")
    IB = islem_sayfasi(wb, "BIST Alım-Satım", "BIST Portföy", "BIST", PB)
    PA = portfoy_sayfasi(wb, "ABD Portföy", "ABD Alım-Satım", "ABD")
    IA = islem_sayfasi(wb, "ABD Alım-Satım", "ABD Portföy", "ABD", PA)
    TB, TA = toplam_adet_sayfasi(wb, IB, IA)
    TC = temettu_sayfasi(wb)
    ozet_sayfasi(ozet, PB, PA, TB, TA, TC)
    kilavuz_sayfasi(wb)

    renkler = {"Özet": "1F3864", "BIST Portföy": "C00000", "BIST Alım-Satım": "E06666",
               "ABD Portföy": "2E75B6", "ABD Alım-Satım": "6FA8DC",
               "Toplam Adet": "7F6000", "Temettü Gelirleri": "38761D", "Nasıl Kullanılır": "808080"}
    for ad, renk in renkler.items():
        wb[ad].sheet_properties.tabColor = renk
    wb.save(cikti)
    print(f"Oluşturuldu: {cikti}")


if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "--test":
        TEST_MODU = True
        olustur(sys.argv[2] if len(sys.argv) > 2 else "test.xlsx")
    else:
        olustur("hisse_takip.xlsx")

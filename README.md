# Hisse Senedi Portföy & Temettü Takip Tablosu

Google E-Tablolar'da çalışan, fiyatları **otomatik güncellenen** portföy ve temettü takip tablosu.

**Dosya:** [`hisse_takip.xlsx`](hisse_takip.xlsx)

## Sekmeler

| Sekme | Ne işe yarar |
|---|---|
| **Özet** | Dolar/TL, BIST 100, S&P 500; BIST + ABD toplam maliyet, değer, kâr/zarar, temettü (₺ ve $) |
| **BIST Portföy** | Borsa İstanbul hisseleri: güncel fiyat, günlük değişim, eldeki adet, ortalama maliyet, kâr/zarar, temettü verimi |
| **BIST Alım-Satım** | Her alım/satım işleminin kaydı |
| **ABD Portföy** | ABD hisseleri, hem **$** hem **₺** (TL maliyeti alım günündeki kurla) |
| **ABD Alım-Satım** | ABD işlemleri; işlem günündeki dolar kuru otomatik gelir |
| **Toplam Adet** | Her hisse için toplam alınan / satılan / eldeki adet, ortalama alış fiyatı, satıştan gerçekleşen kâr |
| **Temettü Gelirleri** | Hesaba yatan temettüler |
| **Nasıl Kullanılır** | Adım adım kullanım kılavuzu |

Popüler temettü hisseleri (BIST ve ABD) örnek olarak hazır gelir; silinebilir veya değiştirilebilir.

## Google E-Tablolar'a yükleme

1. [sheets.google.com](https://sheets.google.com) → **Boş e-tablo**
2. **Dosya → İçe aktar → Yükle** → `hisse_takip.xlsx`
3. **E-tabloyu değiştir** → **Verileri içe aktar**

> Drive'a yükleyip doğrudan açarsanız fiyatlar gelmeyebilir. O durumda **Dosya → Google E-Tablolar olarak kaydet** deyin.

## Notlar

- Fiyatlar `GOOGLEFINANCE` ile gelir. BIST yaklaşık 15 dk gecikmelidir. Anlık alış/satış kademe fiyatları ücretsiz kaynaklarda yoktur.
- Hisse başı temettü tutarı otomatik gelmez, Portföy sekmesine elle girilir.
- Bilgi amaçlıdır, yatırım tavsiyesi değildir.

## Tabloyu yeniden üretmek

```bash
pip install openpyxl
python3 olustur.py            # hisse_takip.xlsx
python3 olustur.py --test x.xlsx   # GOOGLEFINANCE yerine sabit sayılar (formül testi için)
```

---

# Temettü Raporu (BIST, ABD ve temettü fonları)

**Dosya:** [`temettu_raporu.xlsx`](temettu_raporu.xlsx)

| Sekme | Ne işe yarar |
|---|---|
| **Özet** | Son 3/5/7/10 yılın hepsinde temettü ödeyen şirket ve fon sayıları |
| **BIST İlk 20 / ABD İlk 20** | Her yıl temettü ödendi mi (✓/✗) ve **son 3/5/7/10 yılda kaç yıl ödendiği** (ör. 10 yılda 9 → 9). Hepsinde ödediyse yeşil, eksik varsa pembe |
| **ABD Temettü Fonları** | Temettü dağıtan en büyük 20 ABD borsa yatırım fonu (SCHD, VIG, VYM, JEPI…): aynı ✓/✗ ve sayı mantığı, büyüklük, verim, gider oranı, ödeme sıklığı |
| **Batmama Riski (3 yıl)** | 40 şirket için kredi notu + temettü sürekliliğinden basit gösterge |
| **Not Tablosu** | Kredi notu kategorileri |

# Temettü Raporu — Çin (ayrı dosya)

**Dosya:** [`temettu_raporu_cin.xlsx`](temettu_raporu_cin.xlsx)

| Sekme | Ne işe yarar |
|---|---|
| **Çin İlk 20** | Hong Kong / ABD'de işlem gören (Türkiye'den alınabilen) en büyük 20 Çin şirketi: aynı ✓/✗ ve sayı mantığı, Hong Kong kodu, ABD (ADR) kodu, lot büyüklüğü, verim, kredi notu, stopaj notu |
| **Çin Temettü Fonları** | Hong Kong / ABD'de işlem gören, son 3 yılda her yıl dağıtım yapmış 7 Çin temettü fonu: aynı ✓/✗ ve sayı mantığı, büyüklük, verim, gider oranı, lot |
| **Batmama Riski (3 yıl)** | 20 Çin şirketi için aynı gösterge |
| **Not Tablosu** | Kredi notu kategorileri |

İki dosyayı yeniden üretmek için: `python3 temettu_raporu_guncelle.py`
(girdi: `kaynak/temettu_raporu_orijinal.xlsx` + `temettu_fonlari.json` + `temettu_cin.json` + `temettu_cin_fonlari.json`).
Fon ve Çin verisi 11.11.2026 tarihlidir (stockanalysis.com, companiesmarketcap.com, Yahoo Finance, HKEX). Bilgi amaçlıdır, yatırım tavsiyesi değildir.

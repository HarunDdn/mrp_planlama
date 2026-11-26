# -*- coding: utf-8 -*-
"""
MRP girdi sablonunu olusturur.

    python girdi_olustur.py                      -> MRP_Girdi.xlsx (mevcut veri seti)
    python girdi_olustur.py dosya.xlsx           -> baska isimle
    python girdi_olustur.py dosya.xlsx --bos     -> hic veri olmadan bos sablon
    python girdi_olustur.py dosya.xlsx --ornek   -> cok urunlu ornek veri seti

Sablon cok urunlu ve dallanan urun agaclarini destekler; farkli malzemelere
ayri ayri siparis girilebilir, urunler ayni tezgahi paylasabilir.
"""
import datetime as dt
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SON_SATIR = 600                       # acilir listelerin ve bicimin kapsadigi satir sayisi

BASLIK_FILL = PatternFill("solid", fgColor="1F3864")
BASLIK_FONT = Font(color="FFFFFF", bold=True, size=10)
GIRDI_FILL = PatternFill("solid", fgColor="FFF9E6")
NOT_FONT = Font(italic=True, size=9, color="7F7F7F")
BASLIK2 = Font(bold=True, size=12, color="1F3864")
INCE = Side(style="thin", color="D9D9D9")
KENAR = Border(left=INCE, right=INCE, top=INCE, bottom=INCE)


def sayfa(wb, ad, basliklar, satirlar, genislikler, not_metni="", formatlar=None,
          tarih_kolonlari=(), sekme_rengi=None):
    """Not satiri + baslik + veri; asagisi bos ama bicimli."""
    ws = wb.create_sheet(ad) if ad not in wb.sheetnames else wb[ad]
    if sekme_rengi:
        ws.sheet_properties.tabColor = sekme_rengi
    n = len(basliklar)
    if not_metni:
        c = ws.cell(row=1, column=1, value=not_metni)
        c.font = NOT_FONT
        c.alignment = Alignment(vertical="top", wrap_text=True)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
        ws.row_dimensions[1].height = 30
    bs = 2 if not_metni else 1
    for j, b in enumerate(basliklar, start=1):
        c = ws.cell(row=bs, column=j, value=b)
        c.fill = BASLIK_FILL
        c.font = BASLIK_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = KENAR
    ws.row_dimensions[bs].height = 30
    for i, satir in enumerate(satirlar, start=bs + 1):
        for j, v in enumerate(satir, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.border = KENAR
            if formatlar and basliklar[j - 1] in formatlar:
                c.number_format = formatlar[basliklar[j - 1]]
            if j in tarih_kolonlari:
                c.number_format = "DD.MM.YYYY"
    # bos alani bicimle
    for i in range(bs + 1 + len(satirlar), SON_SATIR + 1):
        for j in range(1, n + 1):
            c = ws.cell(row=i, column=j)
            c.border = KENAR
            c.fill = GIRDI_FILL
            if formatlar and basliklar[j - 1] in formatlar:
                c.number_format = formatlar[basliklar[j - 1]]
            if j in tarih_kolonlari:
                c.number_format = "DD.MM.YYYY"
    for j, g in enumerate(genislikler, start=1):
        ws.column_dimensions[get_column_letter(j)].width = g
    ws.freeze_panes = ws.cell(row=bs + 1, column=1)
    ws.auto_filter.ref = "A%d:%s%d" % (bs, get_column_letter(n), SON_SATIR)
    return ws, bs


def liste_dogrulama(ws, kolon_harfi, formul, ilk_satir, mesaj):
    dv = DataValidation(type="list", formula1=formul, allow_blank=True,
                        showErrorMessage=False, promptTitle="Secim", prompt=mesaj)
    ws.add_data_validation(dv)
    dv.add("%s%d:%s%d" % (kolon_harfi, ilk_satir, kolon_harfi, SON_SATIR))


# ====================================================================== #
#  VERI SETLERI
# ====================================================================== #

MEVCUT = {
    "malzemeler": [
        ["MAM012-00012-R00", "610-310-04-140/0 Gegengriff Asien SKT 22 Cm Kpl No:5 (Tencere Sapi 2 BAGALIT)",
         "URETIM", "AD", "MONTAJ", 11.25, 0, 0],
        ["BLZ012-00012-R00", "Bilz. Tak. No:5 610-310-04-140/0 Gegengriff Asien Tenc. Sapi SKT 22 Cm Kpl",
         "URETIM", "AD", "BILEZIK", 7.2, 2734, 0],
        ["CAP012-00012-R00", "Cap. Aln. 610-310-04-140/0 Gegengriff Asien Tenc. Sapi SKT 22 Cm Kpl No:5",
         "URETIM", "AD", "CAPAK", 1.8, 0, 0],
        ["ENJ012-00013-R00", "Enj. Yap. Gegengriff Asien", "URETIM", "AD", "ENJEKSIYON", 136, 11110, 0],
        ["HAM001-00104-R00", "AMITY 101", "SATINALMA", "KG", "", 0, 25982.58406429, 0],
    ],
    "agac": [
        ["MAM012-00012-R00", "BLZ012-00012-R00", 1, 0],
        ["BLZ012-00012-R00", "CAP012-00012-R00", 1, 0],
        ["CAP012-00012-R00", "ENJ012-00013-R00", 1, 0],
        ["ENJ012-00013-R00", "HAM001-00104-R00", 0.078, 0],
    ],
    "seviyeli": [],
    "kaynaklar": [
        ["MONTAJ", "Montaj / komple - 1 kisi, tek vardiya", "KISI", 9, 1, 1, 1, 1, 1, 1, 0, 0],
        ["BILEZIK", "Bilezik takma - 1 kisi, tek vardiya", "KISI", 9, 1, 1, 1, 1, 1, 1, 0, 0],
        ["CAPAK", "Capak alma - 1 kisi, tek vardiya", "KISI", 9, 1, 1, 1, 1, 1, 1, 0, 0],
        ["ENJEKSIYON", "Enjeksiyon - 3 vardiya, molalar dusulmus", "DIGER", 22.5, 1, 1, 1, 1, 1, 1, 0, 0],
    ],
    "rotalar": [],
    "kalipmakine": [],
    "talepler": [
        ["2603010083", 10, "MAM012-00012-R00", 15360, "AD", dt.date(2026, 9, 4), 0],
        ["2603010083", 20, "MAM012-00012-R00", 8000, "AD", dt.date(2026, 9, 11), 0],
        ["2607010851", 10, "MAM012-00012-R00", 19200, "AD", dt.date(2026, 9, 4), 0],
        ["2607010851", 20, "MAM012-00012-R00", 19200, "AD", dt.date(2026, 9, 18), 0],
        ["2607010851", 30, "MAM012-00012-R00", 19200, "AD", dt.date(2026, 9, 25), 0],
        ["2607010851", 40, "MAM012-00012-R00", 19200, "AD", dt.date(2026, 10, 2), 0],
        ["2607010851", 50, "MAM012-00012-R00", 19200, "AD", dt.date(2026, 10, 9), 0],
        ["2607010851", 60, "MAM012-00012-R00", 19200, "AD", dt.date(2026, 10, 16), 0],
    ],
    "acik": [],
}

# ---------------------------------------------------------------------- #
#  COK URUNLU ORNEK - su ozellikleri bir arada gosterir:
#   * 3 nihai urun, farkli malzemelere ayri siparisler
#   * ENJ-500 govdesi HEM URN-100 HEM URN-200 tarafindan kullaniliyor
#   * ENJ-500'un IKI KALIBI var (KALIP-500A oncelikli, dolarsa 500B devreye girer)
#   * Kaliplar PRES-1 / PRES-2 makinelerinde oncelik sirasiyla calisiyor,
#     makineler tum urunler icin ortak
#   * Montajda 1. ve 2. oncelikli kisi var, hizlari farkli
#   * Iki urun agaci formati ayni dosyada
# ---------------------------------------------------------------------- #
ORNEK = {
    "malzemeler": [
        ["URN-100", "Tencere Sapi Komple - Model A", "URETIM", "AD", "", "", 0, 0],
        ["URN-200", "Tencere Sapi Komple - Model B", "URETIM", "AD", "", "", 500, 0],
        ["URN-300", "Kapak Tutamagi Komple", "URETIM", "AD", "", "", 0, 0],
        ["ENJ-500", "Sap Govdesi - Model A ve B'de ORTAK", "URETIM", "AD", "", "", 6000, 0],
        ["ENJ-600", "Tutamak Govdesi", "URETIM", "AD", "", "", 1500, 0],
        ["YRM-900", "Ortak Vida Grubu (3 urunde de kullaniliyor)", "URETIM", "AD", "", "", 8000, 0],
        ["HAM-001", "Bagalit Granul", "SATINALMA", "KG", "", 0, 40000, 500],
        ["HAM-002", "Paslanmaz Tel", "SATINALMA", "KG", "", 0, 3000, 0],
    ],
    # Model A ve B ust-alt cifti formatiyla - ikisi de ayni ENJ-500'u kullaniyor
    "agac": [
        ["URN-100", "ENJ-500", 1, 0],
        ["URN-100", "YRM-900", 2, 0],
        ["URN-200", "ENJ-500", 1, 0],
        ["URN-200", "YRM-900", 2, 0],
        ["ENJ-500", "HAM-001", 0.08, 2],
        ["YRM-900", "HAM-002", 0.005, 0],
    ],
    # URN-300 seviyeli liste formatiyla - iki format ayni dosyada kullanilabilir
    "seviyeli": [
        [0, "URN-300", "", ""],
        [1, "ENJ-600", 1, 0],
        [2, "HAM-001", 0.06, 2],
        [1, "YRM-900", 3, 0],
    ],
    "kaynaklar": [
        ["KISI-1", "Montaj - 1. oncelikli kisi", "KISI", 9, 1, 1, 1, 1, 1, 1, 0, 0],
        ["KISI-2", "Montaj - 2. oncelikli kisi (daha yavas)", "KISI", 9, 1, 1, 1, 1, 1, 1, 0, 0],
        ["KISI-3", "Kapak tutamagi montaji", "KISI", 9, 1, 1, 1, 1, 1, 1, 0, 0],
        ["KISI-4", "Capak / vida grubu", "KISI", 9, 1, 1, 1, 1, 1, 1, 0, 0],
        ["KALIP-500A", "Sap govdesi kalibi A (ana kalip)", "KALIP", 22.5, 1, 1, 1, 1, 1, 1, 1, 0],
        ["KALIP-500B", "Sap govdesi kalibi B (yedek, daha yavas)", "KALIP", 22.5, 1, 1, 1, 1, 1, 1, 1, 0],
        ["KALIP-600A", "Tutamak govdesi kalibi", "KALIP", 22.5, 1, 1, 1, 1, 1, 1, 1, 0],
        ["PRES-1", "Enjeksiyon presi 1 - tum kaliplar icin ortak", "MAKINE", 22.5, 1, 1, 1, 1, 1, 1, 1, 0],
        ["PRES-2", "Enjeksiyon presi 2 - tum kaliplar icin ortak", "MAKINE", 22.5, 1, 1, 1, 1, 1, 1, 1, 0],
    ],
    # Malzeme -> oncelikli kaynak alternatifleri, her biri kendi hiziyla
    "rotalar": [
        ["URN-100", 1, "KISI-1", 12],
        ["URN-100", 2, "KISI-2", 14],
        ["URN-200", 1, "KISI-1", 15],
        ["URN-200", 2, "KISI-2", 17],
        ["URN-300", 1, "KISI-3", 8],
        ["ENJ-500", 1, "KALIP-500A", 136],
        ["ENJ-500", 2, "KALIP-500B", 150],
        ["ENJ-600", 1, "KALIP-600A", 90],
        ["YRM-900", 1, "KISI-4", 4],
    ],
    # Kalip -> oncelikli makineler. Hiz Carpani: 1,10 = o makinede %10 yavas.
    "kalipmakine": [
        ["KALIP-500A", 1, "PRES-1", 1.0],
        ["KALIP-500A", 2, "PRES-2", 1.10],
        ["KALIP-500B", 1, "PRES-2", 1.0],
        ["KALIP-500B", 2, "PRES-1", 1.05],
        ["KALIP-600A", 1, "PRES-2", 1.0],
        ["KALIP-600A", 2, "PRES-1", 1.0],
    ],
    "talepler": [
        ["SIP-0001", 10, "URN-100", 8000, "AD", dt.date(2026, 9, 18), 0],
        ["SIP-0001", 20, "URN-200", 5000, "AD", dt.date(2026, 9, 25), 0],
        ["SIP-0002", 10, "URN-300", 12000, "AD", dt.date(2026, 9, 30), 0],
        ["SIP-0003", 10, "URN-100", 6000, "AD", dt.date(2026, 10, 9), 0],
        ["SIP-0004", 10, "URN-200", 7000, "AD", dt.date(2026, 10, 16), 1],
        ["SIP-0004", 20, "URN-300", 9000, "AD", dt.date(2026, 10, 23), 0],
    ],
    "acik": [
        ["SAS-2026-0410", "HAM-001", 5000, dt.date(2026, 9, 5), "Satin alma siparisi"],
        ["IE-10-20913", "ENJ-500", 1200, dt.date(2026, 8, 24), "Acik is emri"],
        ["REZ-99", "YRM-900", -1500, dt.date(2026, 8, 21), "Baska urun icin rezerve (negatif)"],
    ],
}

BOS = {k: [] for k in MEVCUT}


# ====================================================================== #
def olustur(yol="MRP_Girdi.xlsx", veri=None):
    veri = veri if veri is not None else MEVCUT
    wb = Workbook()
    wb.remove(wb.active)

    # ------------------------------------------------------------------ #
    ws = wb.create_sheet("Kullanim")
    ws.sheet_properties.tabColor = "1F3864"
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 112
    satir = 1

    def yaz(a, b="", kalin=False, baslik=False):
        nonlocal satir
        ca = ws.cell(row=satir, column=1, value=a)
        cb = ws.cell(row=satir, column=2, value=b)
        if baslik:
            ca.font = BASLIK2
        elif kalin:
            ca.font = Font(bold=True, size=10)
        cb.alignment = Alignment(wrap_text=True, vertical="top")
        satir += 1

    yaz("MRP ve PARALEL PLANLAMA - GIRDI DOSYASI", "", baslik=True)
    yaz("")
    yaz("Nasil calistirilir?", "calistir.bat dosyasina cift tikla. Sonuc MRP_Sonuc.xlsx olarak yazilir ve acilir.")
    yaz("")
    yaz("SAYFALAR", "", baslik=True)
    yaz("Malzemeler", "Her malzeme bir satir. Nihai urun, yari mamul ve hammadde ayni listede. "
                      "Tip = URETIM ise Kaynak ve Saniye/Adet zorunlu.")
    yaz("UrunAgaci", "Ust-alt cifti: her satir 'su urunun icine su parcadan su kadar girer'. "
                     "Dallanma ve paylasimli parca serbest.")
    yaz("UrunAgaciSeviyeli", "Ayni bilginin girintili hali: Seviye 0 = nihai urun, 1 = onun alti, 2 = onun alti... "
                             "Hangisi kolayina geliyorsa onu kullan; ikisi ayni dosyada da olabilir.")
    yaz("Kaynaklar", "Kisi / kalip / makine - hepsi ayni listede, Tip kolonuyla ayrilir. "
                     "Gunluk kapasite = Saat/Gun x Paralel Adet. "
                     "BIRDEN FAZLA MALZEME AYNI KAYNAGI KULLANABILIR - kapasite paylasilir.")
    yaz("Rotalar", "Bir malzeme icin birden fazla kaynak: 1. oncelikli kisi/kalip, 2. oncelikli... "
                   "Her alternatifin kendi saniye/adet'i olur. 1. dolarsa kalan ayni gun 2.'ye tasar.")
    yaz("KalipMakine", "Kalip hangi makinelerde, hangi sirayla calisabilir. Uretim kalip VE makine "
                       "kapasitesini ayni anda tuketir; makineler tum urunler icin ortaktir.")
    yaz("Talepler", "Satis siparisleri. FARKLI MALZEMELERE ayri ayri siparis girilebilir.")
    yaz("AcikSiparisler", "Acik is emirleri ve satin alma siparisleri (+), baska urun icin rezervasyon (-).")
    yaz("Parametreler", "Plan baslangici, aktarma kurali, ufuk, tatiller.")
    yaz("KolonEslestirme", "ERP'den yapistirdigin basliklar taninmazsa buradan tanitirsin.")
    yaz("")
    yaz("ERP'DEN KOPYALA-YAPISTIR", "", baslik=True)
    yaz("", "Program baslik satirini kendi bulur; kolon sirasi ve fazladan kolonlar onemli degil. "
            "Bilinen basliklar otomatik taninir (Malzeme, Stok Kodu, Sal. Bel. Num., KalemNo, "
            "Teslim Tarihi, Miktar Br., adet/saniye, ANLIK STOK ...).")
    yaz("", "Taninmayan bir baslik varsa KolonEslestirme sayfasina bir satir ekle: "
            "Sayfa = Talepler, Dosyadaki Baslik = Sevk Tar., Program Alani = Teslim Tarihi.")
    yaz("")
    yaz("PARALEL URETIM", "", baslik=True)
    yaz("", "Alt seviye bir gunluk uretimini tamamladiginda o miktar ERTESI GUN ust seviyenin "
            "kullanimina acilir. Ust seviye alt seviyenin bitmesini beklemez - seviyeler ust uste biner.")
    yaz("", "Parametreler > Aktarma Kurali ile degistirilebilir: "
            "GUNLUK (varsayilan) / ADET (ayni gun akis) / SABIT (belirtilen adet birikince).")
    yaz("")
    yaz("KALIP - MAKINE", "", baslik=True)
    yaz("", "Enjeksiyon adimi kalibi temsil eder. Bir parcanin birden fazla kalibi olabilir "
            "(Rotalar'da oncelik sirasiyla), her kalip birden fazla makinede calisabilir "
            "(KalipMakine'de oncelik sirasiyla). Ayni makineyi farkli urunlerin kaliplari paylasir.")
    yaz("", "Kalip bir makineye baglandiginda orada kalir - gereksiz yere tasinmaz. Sadece o makine "
            "yer vermediginde siradaki makineye gecer. Parametreler > 'Minimum Kalip Kalis (gun)' ile "
            "en az kac gun ayni makinede kalacagini zorlayabilirsin (0 = serbest).")
    yaz("", "Cikti dosyasindaki KalipMakineAtama sayfasi hangi kalibin hangi tarihte hangi makinede "
            "oldugunu ve kac kez tasindigini gosterir.")
    yaz("")
    yaz("SARI HUCRELER", "Veri girilecek alanlar. Istedigin kadar satir ekleyebilirsin; "
                         "program dolu satirlarin hepsini okur.")

    # ------------------------------------------------------------------ #
    ws, bs = sayfa(
        wb, "Malzemeler",
        ["Malzeme Kodu", "Aciklama", "Tip", "Birim", "Kaynak", "Saniye/Adet",
         "Anlik Stok", "Emniyet Stogu"],
        veri["malzemeler"], [24, 58, 12, 8, 14, 12, 15, 14],
        not_metni=("Nihai urun, yari mamul ve hammadde ayni listede.   "
                   "Tip: URETIM = uretilir, SATINALMA = disaridan gelir.   "
                   "Kaynak + Saniye/Adet: malzemenin TEK kaynagi varsa buraya yaz. "
                   "BIRDEN FAZLA kaynak/kalip varsa bu iki kolonu bos birak ve Rotalar sayfasini doldur.   "
                   "Emniyet Stogu bos birakilabilir."),
        formatlar={"Saniye/Adet": "#,##0.000", "Anlik Stok": "#,##0.000",
                   "Emniyet Stogu": "#,##0.000"},
        sekme_rengi="2E75B6")
    liste_dogrulama(ws, "C", '"URETIM,SATINALMA"', bs + 1, "URETIM veya SATINALMA")
    liste_dogrulama(ws, "E", "=Kaynaklar!$A$3:$A$%d" % SON_SATIR, bs + 1,
                    "Tek kaynak icin; birden fazla kaynak varsa Rotalar sayfasini kullan")
    MAL_ARALIK = "=Malzemeler!$A$%d:$A$%d" % (bs + 1, SON_SATIR)

    # ------------------------------------------------------------------ #
    ws, bs2 = sayfa(
        wb, "UrunAgaci",
        ["Ust Malzeme", "Alt Malzeme", "Miktar (alt/1 ust)", "Fire %"],
        veri["agac"], [26, 26, 20, 10],
        not_metni=("Her satir bir baglanti: 1 adet Ust Malzeme icin kac Alt Malzeme gerekir.   "
                   "Ayni ust icin birden fazla alt satir yazilabilir (dallanma).   "
                   "Ayni alt parca birden fazla urunde kullanilabilir (paylasimli parca)."),
        formatlar={"Miktar (alt/1 ust)": "#,##0.000000", "Fire %": "0.00"},
        sekme_rengi="2E75B6")
    liste_dogrulama(ws, "A", MAL_ARALIK, bs2 + 1, "Malzemeler sayfasindan")
    liste_dogrulama(ws, "B", MAL_ARALIK, bs2 + 1, "Malzemeler sayfasindan")

    # ------------------------------------------------------------------ #
    ws, bs3 = sayfa(
        wb, "UrunAgaciSeviyeli",
        ["Seviye", "Malzeme", "Miktar (alt/1 ust)", "Fire %"],
        veri["seviyeli"], [10, 26, 20, 10],
        not_metni=("UrunAgaci sayfasinin alternatifi - hangisi kolayina geliyorsa onu doldur, "
                   "ikisi ayni dosyada da kullanilabilir.   Seviye 0 = nihai urun (Miktar bos), "
                   "1 = onun alti, 2 = onun alti...   Bir satirin ustu, yukarisindaki en yakin "
                   "(seviye-1) satiridir.   Ornek:  0 URN-300 |  1 YRM-310 (1 adet) |  2 HAM-001 (0,06 kg) |  1 YRM-900 (3 adet)"),
        formatlar={"Seviye": "0", "Miktar (alt/1 ust)": "#,##0.000000", "Fire %": "0.00"},
        sekme_rengi="2E75B6")
    liste_dogrulama(ws, "B", MAL_ARALIK, bs3 + 1, "Malzemeler sayfasindan")

    # ------------------------------------------------------------------ #
    ws, bs4 = sayfa(
        wb, "Kaynaklar",
        ["Kaynak Kodu", "Aciklama", "Tip", "Saat/Gun", "Paralel Adet",
         "Pzt", "Sal", "Car", "Per", "Cum", "Cmt", "Paz"],
        veri["kaynaklar"], [16, 46, 11, 11, 13, 6, 6, 6, 6, 6, 6, 6],
        not_metni=("Kisi, kalip ve makine ayni listede - Tip kolonuyla ayrilir.   "
                   "Saat/Gun = bir gunde fiilen calisilan net saat (molalar dusulmus).   "
                   "Paralel Adet = birebir ayni isi yapan adet sayisi.   "
                   "Gunluk kapasite = Saat/Gun x Paralel Adet.   Gun kolonlari: 1 = calisilir, 0 = calisilmaz.   "
                   "Birden fazla malzeme ayni kaynagi gosterirse kapasiteyi PAYLASIRLAR."),
        formatlar={"Saat/Gun": "#,##0.00", "Paralel Adet": "#,##0"},
        sekme_rengi="C55A11")
    liste_dogrulama(ws, "C", '"KISI,KALIP,MAKINE,DIGER"', bs4 + 1,
                    "KALIP ise KalipMakine sayfasinda makineye baglanmali")
    KAYNAK_ARALIK = "=Kaynaklar!$A$%d:$A$%d" % (bs4 + 1, SON_SATIR)

    # ------------------------------------------------------------------ #
    ws, bsr = sayfa(
        wb, "Rotalar",
        ["Malzeme", "Oncelik", "Kaynak", "Saniye/Adet"],
        veri["rotalar"], [26, 10, 20, 14],
        not_metni=("BIR MALZEME ICIN BIRDEN FAZLA KAYNAK. Her satir bir alternatif.   "
                   "Oncelik 1 = once bu kaynak calisir; kapasitesi o gun dolarsa KALAN MIKTAR "
                   "ayni gun oncelik 2'ye tasar.   Saniye/Adet her alternatif icin ayri girilir "
                   "(ayni parca farkli kiside/kalipta farkli hizda uretilebilir).   "
                   "Bu sayfada yer almayan malzemeler Malzemeler sayfasindaki tek kaynagi kullanir."),
        formatlar={"Oncelik": "0", "Saniye/Adet": "#,##0.000"},
        sekme_rengi="C55A11")
    liste_dogrulama(ws, "A", MAL_ARALIK, bsr + 1, "Malzemeler sayfasindan")
    liste_dogrulama(ws, "C", KAYNAK_ARALIK, bsr + 1, "Kaynaklar sayfasindan (kisi veya kalip)")

    # ------------------------------------------------------------------ #
    ws, bsk = sayfa(
        wb, "KalipMakine",
        ["Kalip", "Oncelik", "Makine", "Hiz Carpani"],
        veri["kalipmakine"], [22, 10, 20, 14],
        not_metni=("Bir KALIP hangi makinelerde calisabilir, hangi sirayla.   "
                   "Uretim hem kalibin hem makinenin kapasitesini AYNI ANDA tuketir; "
                   "makineler tum urunler icin ortaktir.   "
                   "Hiz Carpani: 1 = normal, 1,10 = o makinede %10 yavas (sure = Saniye/Adet x Carpan).   "
                   "Kalip bir kez bir makineye baglandiginda orada kalir; ancak o makine yer "
                   "vermezse siradaki makineye tasinir."),
        formatlar={"Oncelik": "0", "Hiz Carpani": "#,##0.000"},
        sekme_rengi="C55A11")
    liste_dogrulama(ws, "A", KAYNAK_ARALIK, bsk + 1, "Tip = KALIP olan kaynak")
    liste_dogrulama(ws, "C", KAYNAK_ARALIK, bsk + 1, "Tip = MAKINE olan kaynak")

    # ------------------------------------------------------------------ #
    ws, bs5 = sayfa(
        wb, "Talepler",
        ["Belge No", "Kalem", "Malzeme", "Miktar", "Birim", "Teslim Tarihi", "Oncelik"],
        veri["talepler"], [18, 8, 26, 14, 8, 15, 10],
        not_metni=("Satis siparisleri. Farkli malzemelere ayri ayri siparis girilebilir; "
                   "ayni malzemenin birden fazla teslim kalemi olabilir.   "
                   "Oncelik: 0 = normal, kucuk sayi once planlanir. Bos birakilirsa teslim tarihine gore siralanir."),
        formatlar={"Miktar": "#,##0.000", "Oncelik": "0"},
        tarih_kolonlari=(6,), sekme_rengi="548235")
    liste_dogrulama(ws, "C", MAL_ARALIK, bs5 + 1, "Malzemeler sayfasindan")

    # ------------------------------------------------------------------ #
    ws, bs6 = sayfa(
        wb, "AcikSiparisler",
        ["Belge", "Malzeme", "Miktar", "Tarih", "Aciklama"],
        veri["acik"], [30, 26, 14, 14, 46],
        not_metni=("Acik is emirleri ve satin alma siparisleri.   "
                   "Miktar POZITIF = stoga giris (uretim tamamlanacak / mal gelecek).   "
                   "Miktar NEGATIF = cikis (baska urunun tuketimi, rezervasyon).   "
                   "Gecmis tarihli girisler plan baslangic gunune cekilir."),
        formatlar={"Miktar": "#,##0.000"},
        tarih_kolonlari=(4,), sekme_rengi="548235")
    liste_dogrulama(ws, "B", MAL_ARALIK, bs6 + 1, "Malzemeler sayfasindan")

    # ------------------------------------------------------------------ #
    par = [
        ["Plan Baslangic Tarihi", dt.date(2026, 8, 19),
         "Ileri planlamanin basladigi gun. Bos birakilirsa bilgisayarin tarihi kullanilir."],
        ["Aktarma Kurali", "GUNLUK",
         "GUNLUK = alt seviyenin 1 gunluk uretimi birikince ust seviye baslar. "
         "ADET = ayni gun icinde de akis olur. SABIT = asagidaki parti kadar birikince."],
        ["Aktarma Partisi", 0, "Aktarma Kurali = SABIT ise gecerli (adet)."],
        ["Planlama Ufku (gun)", 730, "Ileri simulasyonun kac gun surecegi. Kapasite yetersizse buyuk tutulmali."],
        ["Siralama Kurali", "EDD", "EDD = en erken teslim tarihi once. FIFO = belge sirasina gore."],
        ["Tatiller", "", "Calisilmayan gunler, noktali virgulle: 30.08.2026;29.10.2026"],
        ["Senaryo Analizi", "EVET",
         "EVET ise darbogaza kademeli tezgah eklenip 'kac tezgah gerekir' sorusu cevaplanir."],
        ["Minimum Kalip Kalis (gun)", 0,
         "Bir kalip bir makineye baglandiktan sonra en az kac gun orada kalsin. "
         "0 = serbest (yer bulamazsa hemen tasinir). 7 = bir hafta ayni makinede kalir."],
        ["Satinalma Kisiti", "HAYIR",
         "HAYIR = satin alinan malzeme her zaman bulunur varsayilir, acik raporlanir. "
         "EVET = stok bitince uretim durur."],
    ]
    ws, bs7 = sayfa(wb, "Parametreler", ["Parametre", "Deger", "Aciklama"], par,
                    [26, 20, 92], not_metni="Parametre isimlerini degistirme; sadece Deger kolonunu duzenle.",
                    sekme_rengi="7F7F7F")
    ws.auto_filter.ref = None
    ws.cell(row=bs7 + 1, column=2).number_format = "DD.MM.YYYY"
    for i in range(bs7 + 1, bs7 + 1 + len(par)):
        ws.cell(row=i, column=1).font = Font(bold=True, size=10)
        ws.cell(row=i, column=2).fill = GIRDI_FILL
        ws.cell(row=i, column=3).alignment = Alignment(wrap_text=True, vertical="top")
    for i in range(bs7 + 1 + len(par), SON_SATIR + 1):
        for j in range(1, 4):
            ws.cell(row=i, column=j).fill = PatternFill()
            ws.cell(row=i, column=j).border = Border()

    # ------------------------------------------------------------------ #
    alanlar = ("Malzeme Kodu,Aciklama,Tip,Birim,Kaynak,Saniye/Adet,Anlik Stok,Emniyet Stogu,"
               "Ust Malzeme,Alt Malzeme,Miktar (alt/1 ust),Fire %,Seviye,Saat/Gun,Paralel Adet,"
               "Belge No,Kalem,Miktar,Teslim Tarihi,Oncelik")
    ws, bs8 = sayfa(
        wb, "KolonEslestirme", ["Sayfa", "Dosyadaki Baslik", "Program Alani"],
        [], [22, 34, 30],
        not_metni=("ERP'den yapistirdigin bir baslik program tarafindan taninmazsa buraya yaz.   "
                   "Ornek:  Sayfa = Talepler | Dosyadaki Baslik = Sevk Tar. | Program Alani = Teslim Tarihi   "
                   "Bilinen basliklar (Malzeme, Stok Kodu, Sal. Bel. Num., KalemNo, Miktar Br., "
                   "adet/saniye, ANLIK STOK ...) zaten otomatik taninir, buraya yazmana gerek yok."),
        sekme_rengi="7F7F7F")
    liste_dogrulama(ws, "A", '"Malzemeler,UrunAgaci,UrunAgaciSeviyeli,Kaynaklar,Talepler,AcikSiparisler"',
                    bs8 + 1, "Hangi sayfadaki baslik?")
    liste_dogrulama(ws, "C", '"%s"' % alanlar, bs8 + 1, "Program alani")

    wb.save(yol)
    return yol


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    bayrak = [a for a in sys.argv[1:] if a.startswith("--")]
    yol = args[0] if args else "MRP_Girdi.xlsx"
    veri = BOS if "--bos" in bayrak else (ORNEK if "--ornek" in bayrak else MEVCUT)
    print("Olusturuldu:", olustur(yol, veri))

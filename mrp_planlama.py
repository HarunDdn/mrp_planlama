# -*- coding: utf-8 -*-
"""
MRP + PARALEL (LOT STREAMING) SONLU KAPASITE PLANLAMA
=====================================================

Kullanim:
    python mrp_planlama.py MRP_Girdi.xlsx MRP_Sonuc.xlsx

Mantik
------
1) MRP PATLATMA : Cok seviyeli urun agaci patlatilir, her seviyede anlik stok
                  ve acik siparisler netlenir -> planli emir (is) listesi.

2) GERI PLANLAMA: Her talep, teslim tarihinden geriye dogru sonlu kapasiteyle
                  yuklenir. Ust seviye hangi gun ne kadar uretiyorsa, alt seviye
                  o miktari BIR GUN ONCESINE kadar hazirlamak zorundadir.
                  => Seviyeler ust uste biner (paralel calisir), "alt seviye
                     tamamen bitmeden ust seviye baslayamaz" kisiti YOKTUR.

3) ILERI PLANLAMA: Plan baslangic gununden itibaren gun gun simulasyon.
                  Her kaynak o gunku kapasitesi kadar uretir; bir gunde uretilen
                  miktar ERTESI GUN ust seviyenin kullanimina acilir.
                  Gercekci bitis tarihleri ve gecikmeler buradan cikar.

Aktarma kurali (Parametreler sayfasi):
    GUNLUK -> alt seviyenin 1 gunluk uretimi birikince ust seviye baslar (varsayilan)
    ADET   -> ayni gun icinde de akis olur (gecikme yok)
    SABIT  -> belirtilen adet birikince
"""

import datetime as dt
import math
import sys
from collections import defaultdict, OrderedDict

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule

EPS = 1e-6


# ====================================================================== #
#  YARDIMCILAR
# ====================================================================== #

def _s(v):
    return "" if v is None else str(v).strip()


def _f(v, vars=0.0):
    if v is None or (isinstance(v, str) and not v.strip()):
        return vars
    try:
        return float(v)
    except (TypeError, ValueError):
        return vars


def _d(v):
    """Hucreyi tarihe cevir."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError("Tarih okunamadi: %r" % v)


# --------------------------------------------------------------------- #
#  ERP BASLIK TAKMA ADLARI
#  Farkli ERP ciktilarindan kopyala-yapistir yapilabilsin diye her mantiksal
#  alan icin bilinen basliklar. Buraya uymayan basliklar KolonEslestirme
#  sayfasindan tanitilabilir.
# --------------------------------------------------------------------- #
ALANLAR = {
    "malzeme_kodu": ["Malzeme Kodu", "Malzeme", "Stok Kodu", "Malzeme No", "Stok No",
                     "Urun Agaci", "Kod", "Kalem Kodu", "Parca No", "Item", "Material"],
    "aciklama":     ["Aciklama", "Malzeme Aciklamasi", "Tanim", "Malzeme Adi", "Stok Adi",
                     "Aciklamasi", "Description"],
    "tip":          ["Tip", "Malzeme Tipi", "Temin Sekli", "Tedarik Tipi"],
    "birim":        ["Birim", "Miktar Br.", "Br.", "Olcu Birimi", "Birimi", "UOM"],
    # NOT: "Makine" ve "Tezgah" bilerek burada degil - KalipMakine sayfasindaki
    # makine kolonuyla cakisiyordu.
    "kaynak":       ["Kaynak", "Kaynak Kodu", "Is Merkezi", "Operasyon", "Is Istasyonu",
                     "Kaynak/Kalip"],
    "saniye_adet":  ["Saniye/Adet", "adet/saniye", "Sn/Adet", "Cevrim Suresi", "Birim Sure",
                     "Sure (sn)", "Uretim Suresi", "Birim Zaman"],
    "stok":         ["Anlik Stok", "Stok", "ANLIK STOK", "Eldeki Stok", "Mevcut Stok",
                     "Stok Miktari", "Baslangic Stogu"],
    "emniyet":      ["Emniyet Stogu", "Guvenlik Stogu", "Min Stok", "Emniyet"],
    "ust_malzeme":  ["Ust Malzeme", "Ust", "Ana Malzeme", "Ust Urun", "Parent"],
    "alt_malzeme":  ["Alt Malzeme", "Alt", "Bilesenler", "Bilesen", "Component", "Child"],
    # NOT: sade "Miktar" bilerek burada degil - Talepler sayfasindaki miktarla
    # cakisiyordu. Urun agacinda "Miktar" basligi varsa asagidaki okuyucu
    # once "oran"a, bulamazsa "miktar"a duser.
    "oran":         ["Miktar (alt/1 ust)", "Oran", "Sarf Miktari", "Birim Sarf",
                     "Kullanim Miktari", "Katsayi", "Miktar (alt/1ust)"],
    "fire":         ["Fire %", "Fire", "Kayip %", "Iskarta %"],
    "seviye":       ["Seviye", "Kademe", "Level", "Duzey"],
    "kalip":        ["Kalip", "Kalip Kodu", "Mold"],
    "makine":       ["Makine", "Makine Kodu", "Pres", "Tezgah Kodu", "Machine"],
    "hiz_carpani":  ["Hiz Carpani", "Hiz Katsayisi", "Sure Carpani", "Carpan"],
    # Bir kaynak calisirken ayni anda tuketilen ikinci kaynak (orn. makineyi
    # calistiran operator havuzu). Katsayi 1 = 1 saat makine -> 1 saat operator.
    "ek_kaynak":    ["Ek Kaynak", "Operator Havuzu", "Yardimci Kaynak", "Operator"],
    "ek_katsayi":   ["Ek Kaynak Katsayisi", "Ek Kaynak Carpani", "Operator/Makine"],
    "kapasite_grubu": ["Kapasite Grubu", "Is Merkezi / Kapasite Grubu", "Kapasite Gurubu"],
    # Kaynak kapatmalari (izin / ariza / bakim)
    "kapanis_bas":  ["Baslangic", "Baslangic Tarihi", "Bas", "Ilk Gun"],
    "kapanis_bit":  ["Bitis", "Bitis Tarihi", "Bit", "Son Gun"],
    "kapanis_turu": ["Tur", "Sebep", "Neden", "Kapanis Turu"],
    "kapanis_orani": ["Kalan Kapasite %", "Kapasite %", "Oran %"],
    # Bir isi AYNI ANDA yapmasi gereken kisi sayisi (ERP: CALISAN PERSONEL SAYISI).
    # Saniye/Adet bu ekibin 1 adedi bitirme suresidir.
    "gerekli_kisi": ["Gerekli Kisi", "Calisan Personel Sayisi", "Ekip Buyuklugu",
                     "Gerekli Personel"],
    # Partili (batch) calisan kaynaklar, orn. firin: Q adet T saatte
    "parti_miktari": ["Parti Miktari", "Batch Miktari", "Parti Adedi"],
    "parti_saat":    ["Parti Suresi (saat)", "Batch Suresi (saat)", "Parti Suresi"],
    "saat_gun":     ["Saat/Gun", "Gunluk Saat", "Calisma Saati", "Kapasite (saat)"],
    "paralel":      ["Paralel Adet", "Paralel", "Tezgah Sayisi", "Makine Sayisi", "Kisi Sayisi"],
    "belge":        ["Belge No", "Belge", "Sal. Bel. Num.", "Satis Belge No", "Siparis No",
                     "Sip. No", "Belge Numarasi", "Evrak No"],
    "kalem":        ["Kalem", "KalemNo", "Kalem No", "Satir", "Sira", "Kalem Numarasi"],
    "miktar":       ["Miktar", "Siparis Miktari", "Talep Miktari", "Adet", "Sip. Miktar"],
    "tarih":        ["Teslim Tarihi", "Tarih", "Termin", "Termin Tarihi", "Ihtiyac Tarihi",
                     "Sevk Tarihi", "Planlanan Tarih"],
    "oncelik":      ["Oncelik", "Priority", "Aciliyet"],
}


def _norm(x):
    x = _s(x).lower()
    for a, b in (("ı", "i"), ("İ", "i"), ("ş", "s"), ("ğ", "g"), ("ü", "u"),
                 ("ö", "o"), ("ç", "c"), (" ", ""), ("_", ""), ("-", ""),
                 (".", ""), ("/", ""), ("(", ""), (")", ""), ("%", "")):
        x = x.replace(a, b)
    return x


_ALIAS = {}
for _alan, _adlar in ALANLAR.items():
    for _ad in _adlar:
        _ALIAS.setdefault(_norm(_ad), _alan)


def al(satir, alan, vars=None):
    """Satirdan mantiksal alani cek (baslik takma adlarina toleransli)."""
    for k, v in satir.items():
        if _ALIAS.get(_norm(k)) == alan:
            return v
    return vars


OKUNAN_BASLIKLAR = {}          # sayfa -> okunan basliklar (hata mesajlari icin)


def sayfa_oku(wb, ad, eslestirme=None):
    """Basligi otomatik bulur, sozluk listesi dondurur."""
    if ad not in wb.sheetnames:
        return []
    ws = wb[ad]
    baslik_satiri = None
    basliklar = []
    for r in range(1, min(ws.max_row, 10) + 1):
        vals = [_s(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        dolu = [v for v in vals if v]
        if len(dolu) >= 2:
            baslik_satiri = r
            basliklar = vals
            break
    if baslik_satiri is None:
        return []
    # kullanicinin tanittigi baslik eslestirmeleri (KolonEslestirme sayfasi)
    if eslestirme:
        basliklar = [eslestirme.get((_norm(ad), _norm(b)), b) for b in basliklar]
    OKUNAN_BASLIKLAR[ad] = [b for b in basliklar if b]

    out = []
    for r in range(baslik_satiri + 1, ws.max_row + 1):
        satir = {}
        bos = True
        for c, b in enumerate(basliklar, start=1):
            if not b:
                continue
            v = ws.cell(row=r, column=c).value
            satir[b] = v
            if v is not None and _s(v) != "":
                bos = False
        if not bos:
            out.append(satir)
    return out


# ====================================================================== #
#  VERI YAPILARI
# ====================================================================== #

class Kaynak:
    def __init__(self, kod, aciklama, tip, saat_gun, paralel, gunler, tatiller,
                 ek_kaynak="", ek_katsayi=1.0):
        self.kod = kod
        self.aciklama = aciklama
        self.tip = (tip or "").upper()        # KISI / KALIP / MAKINE / DIGER
        self.saat_gun = saat_gun
        self.paralel = max(1, int(paralel or 1))
        self.gunler = gunler                 # [Pzt..Paz] 1/0
        self.tatiller = tatiller
        self.makineler = []                  # [(makine_kodu, hiz_carpani)] oncelik sirasiyla
        self.ek_kaynak = ek_kaynak           # ayni anda tuketilen ikinci kaynak
        self.ek_katsayi = ek_katsayi if ek_katsayi > 0 else 1.0
        # Tarih araligi bazli kapatmalar: izin, ariza, bakim...
        # [(baslangic, bitis, kalan_kapasite_orani 0..1, tur, not)]
        self.kapanis = []

    def kapanis_orani(self, g):
        """O gun icin kalan kapasite orani (1 = tam acik, 0 = tamamen kapali)."""
        o = 1.0
        for bas, bit, oran, _t, _n in self.kapanis:
            if bas <= g <= bit:
                o = min(o, oran)
        return o

    def kapanis_nedeni(self, g):
        for bas, bit, oran, tur, nt in self.kapanis:
            if bas <= g <= bit:
                return (tur or "KAPALI") + (" - " + nt if nt else "")
        return ""

    def calisir(self, g):
        if g in self.tatiller:
            return False
        if self.gunler[g.weekday()] != 1:
            return False
        return self.kapanis_orani(g) > 1e-9

    def kapasite_sn(self, g):
        if not self.calisir(g):
            return 0.0
        return self.saat_gun * 3600.0 * self.paralel * self.kapanis_orani(g)


class Alternatif:
    """
    Bir malzemenin bir uretim secenegi: hangi kaynakta (kisi / kalip),
    adet basina kac saniyede. Kaynak bir KALIP ise ayrica bir MAKINE'ye
    baglanmasi gerekir; makine listesi Kaynak.makineler'de oncelik sirasiyla.
    """
    def __init__(self, kaynak, sn, oncelik):
        self.kaynak = kaynak
        self.sn = sn
        self.oncelik = oncelik

    def __repr__(self):
        return "<Alt %s %.3f sn (oncelik %g)>" % (self.kaynak, self.sn, self.oncelik)


class Malzeme:
    def __init__(self, kod, aciklama, tip, birim, kaynak, sn, stok, emniyet):
        self.kod = kod
        self.aciklama = aciklama
        self.tip = tip                       # URETIM / SATINALMA
        self.birim = birim
        self.kaynak = kaynak                 # varsayilan / birincil kaynak
        self.sn = sn                         # saniye / adet
        self.stok = stok
        self.emniyet = emniyet
        self.llc = 0                         # low level code
        self.cocuklar = []                   # [(kod, oran)]
        self.ebeveynler = []
        self.alternatifler = []              # [Alternatif] oncelik sirasiyla
        self.grup = ""                       # kapasite grubu (ham is merkezi adi)
        self.gerekli_kisi = 1                # isi ayni anda yapmasi gereken kisi sayisi
        self.parti = None                    # (parti adedi, parti suresi sn) - firin vb.
        self.ekip = 1                        # fiilen rezerve edilebilen kisi (<= gerekli_kisi)
        self.ekip_carpan = 1.0               # ekip kucukse sure bu kadar uzar (adam-saat korunur)

    @property
    def kaynaklar_metni(self):
        return " > ".join(a.kaynak for a in self.alternatifler)


class Is:
    """Planli emir."""
    _sayac = 0

    def __init__(self, malzeme, miktar, ihtiyac_tarihi, kaynak_talep, seviye, oncelik=0.0):
        Is._sayac += 1
        self.no = Is._sayac
        self.malzeme = malzeme
        self.miktar = miktar
        self.ihtiyac_tarihi = ihtiyac_tarihi
        self.kaynak_talep = kaynak_talep     # tetikleyen satis siparisi etiketi
        self.seviye = seviye
        self.oncelik = oncelik               # tetikleyen talebin onceligi
        self.kalan = miktar
        self.gunluk = defaultdict(float)     # ileri plan uretimi
        self.baslangic = None
        self.bitis = None


# ====================================================================== #
#  MODEL
# ====================================================================== #

class Model:
    def __init__(self, girdi):
        wb = load_workbook(girdi, data_only=True)

        # --- parametreler ---
        # --- kullanicinin tanittigi baslik eslestirmeleri ---
        # KolonEslestirme sayfasi: Sayfa | Dosyadaki Baslik | Program Alani
        self.eslestirme = {}
        for satir in sayfa_oku(wb, "KolonEslestirme"):
            sayfa = _s(satir.get("Sayfa") or satir.get("sayfa"))
            eski = _s(satir.get("Dosyadaki Baslik") or satir.get("Dosyadaki Başlık"))
            yeni = _s(satir.get("Program Alani") or satir.get("Program Alanı"))
            if sayfa and eski and yeni:
                self.eslestirme[(_norm(sayfa), _norm(eski))] = yeni
        esl = self.eslestirme

        p = {}
        for satir in sayfa_oku(wb, "Parametreler", esl):
            ad = _s(satir.get("Parametre"))
            if ad:
                p[ad.lower()] = satir.get("Deger")
        self.p = p

        self.tatiller = set()
        for parca in _s(p.get("tatiller", "")).replace(",", ";").split(";"):
            if parca.strip():
                self.tatiller.add(_d(parca.strip()))

        self.baslangic = _d(p.get("plan baslangic tarihi")) or dt.date.today()
        self.ufuk = int(_f(p.get("planlama ufku (gun)"), 730))
        self.aktarma_kurali = (_s(p.get("aktarma kurali")) or "GUNLUK").upper()
        self.aktarma_partisi = _f(p.get("aktarma partisi"), 0)
        self.siralama = (_s(p.get("siralama kurali")) or "EDD").upper()
        self.satinalma_kisit = (_s(p.get("satinalma kisiti")) or "HAYIR").upper() == "EVET"
        self.min_kalip_kalis = int(_f(p.get("minimum kalip kalis (gun)"), 0))
        # Bir kaynagin ayni gun icinde kac FARKLI urune gecebilecegi.
        # 1 = gun icinde tek is; 0 = sinirsiz. Gecikmis isler bu kurali asabilir.
        self.gunluk_is_cesidi = int(_f(p.get("gunluk is cesidi"), 0))
        # Bir isin ayni gun kac FARKLI kaynaga (kisiye / ekibe) bolunebilecegi.
        # 1 = is tek kisiye verilir, erken bitsin diye baskalarina bolunmez;
        # 0 = sinirsiz. Teslim tarihi gelmis/gecmis isler bu kurali asabilir.
        self.is_basina_kaynak = int(_f(p.get("is basina azami kaynak"), 0))
        # Takviye hesabinda birakilan emniyet payi (is gunu). Kaynak gun icinde
        # baska isler de aldigi icin nominal hesap iyimser kalir; bu pay kadar
        # erken bitirecek sekilde takviye hesaplanir.
        self.takviye_payi = int(_f(p.get("takviye emniyet payi (gun)"), 2))
        # Mal erken hazir olsa bile teslim tarihinden once sevk edilsin mi?
        # HAYIR = stokta bekler, cikis teslim haftasinda gorunur (onerilen).
        self.erken_sevkiyat = (_s(p.get("erken sevkiyat")) or "HAYIR").upper() == "EVET"

        # --- kaynaklar ---
        self.kaynaklar = {}
        for satir in sayfa_oku(wb, "Kaynaklar", esl):
            kod = _s(al(satir, "kaynak"))
            if not kod:
                continue
            gunler = [int(_f(satir.get(g), 0)) for g in
                      ("Pzt", "Sal", "Car", "Per", "Cum", "Cmt", "Paz")]
            if not any(gunler):
                gunler = [1, 1, 1, 1, 1, 0, 0]
            self.kaynaklar[kod] = Kaynak(
                kod, _s(al(satir, "aciklama")),
                _s(al(satir, "tip")),
                _f(al(satir, "saat_gun"), 8.0),
                _f(al(satir, "paralel"), 1),
                gunler, self.tatiller,
                _s(al(satir, "ek_kaynak")),
                _f(al(satir, "ek_katsayi"), 1.0))

        # --- kaynak kapatmalari (izin / ariza / bakim) ---
        self.kapanis_sayisi = 0
        for satir in sayfa_oku(wb, "Kapanislar", esl):
            kod = _s(al(satir, "kaynak"))
            if not kod:
                continue
            k = self.kaynaklar.get(kod)
            if k is None:
                self.uyarilar.append(
                    "Kapanislar: '%s' kaynagi tanimli degil, satir atlandi." % kod)
                continue
            bas = _d(al(satir, "kapanis_bas"))
            bit = _d(al(satir, "kapanis_bit")) or bas
            if bas is None:
                continue
            if bit < bas:
                bas, bit = bit, bas
            oran = _f(al(satir, "kapanis_orani"), 0.0)
            oran = 0.0 if oran is None else max(0.0, min(100.0, float(oran))) / 100.0
            k.kapanis.append((bas, bit, oran,
                              _s(al(satir, "kapanis_turu")) or "KAPALI",
                              _s(al(satir, "aciklama"))))
            self.kapanis_sayisi += 1

        # --- kalip -> makine (oncelik sirasiyla) ---
        km = defaultdict(list)
        for satir in sayfa_oku(wb, "KalipMakine", esl):
            kalip = _s(al(satir, "kalip"))
            makine = _s(al(satir, "makine"))
            if not kalip or not makine:
                continue
            for k in (kalip, makine):
                if k not in self.kaynaklar:
                    raise ValueError("KalipMakine: '%s' Kaynaklar sayfasinda tanimli degil." % k)
            km[kalip].append((_f(al(satir, "oncelik"), 0), makine,
                              _f(al(satir, "hiz_carpani"), 1.0) or 1.0))
        for kalip, lst in km.items():
            lst.sort(key=lambda x: x[0])
            self.kaynaklar[kalip].makineler = [(m, c) for _, m, c in lst]

        # --- malzemeler ---
        self.malzemeler = OrderedDict()
        self.uyarilar = []
        ham_malzeme = sayfa_oku(wb, "Malzemeler", esl)
        for satir in ham_malzeme:
            kod = _s(al(satir, "malzeme_kodu"))
            if not kod:
                continue
            if kod in self.malzemeler:
                self.uyarilar.append("Malzemeler: '%s' birden fazla kez tanimli, "
                                     "ilki kullanildi." % kod)
                continue
            self.malzemeler[kod] = Malzeme(
                kod,
                _s(al(satir, "aciklama")),
                (_s(al(satir, "tip")) or "URETIM").upper(),
                _s(al(satir, "birim")) or "AD",
                _s(al(satir, "kaynak")),
                _f(al(satir, "saniye_adet")),
                _f(al(satir, "stok")),
                _f(al(satir, "emniyet")))
            self.malzemeler[kod].grup = _s(al(satir, "kapasite_grubu"))

        self._baslik_kontrol("Malzemeler", ham_malzeme, self.malzemeler, "Malzeme Kodu")

        # --- rotalar: malzeme -> oncelikli kaynak alternatifleri ---
        rota = defaultdict(list)
        for satir in sayfa_oku(wb, "Rotalar", esl):
            mkod = _s(al(satir, "malzeme_kodu"))
            kkod = _s(al(satir, "kaynak"))
            if not mkod or not kkod:
                continue
            if mkod not in self.malzemeler:
                raise ValueError("Rotalar: '%s' Malzemeler sayfasinda tanimli degil." % mkod)
            if kkod not in self.kaynaklar:
                raise ValueError("Rotalar: '%s' kaynagi Kaynaklar sayfasinda tanimli degil "
                                 "(malzeme %s)." % (kkod, mkod))
            sn = _f(al(satir, "saniye_adet"), 0.0) or self.malzemeler[mkod].sn
            gk = int(_f(al(satir, "gerekli_kisi"), 1) or 1)
            pq = _f(al(satir, "parti_miktari"), 0.0) or 0.0
            pt = _f(al(satir, "parti_saat"), 0.0) or 0.0
            if pq > 1 and pt > 0:
                self.malzemeler[mkod].parti = (pq, pt * 3600.0)
            if gk > 1:
                self.malzemeler[mkod].gerekli_kisi = max(
                    self.malzemeler[mkod].gerekli_kisi, gk)
            rota[mkod].append(Alternatif(kkod, sn, _f(al(satir, "oncelik"), 0)))
        for mkod, lst in rota.items():
            lst.sort(key=lambda a: a.oncelik)
            self.malzemeler[mkod].alternatifler = lst
        # Rotalar'da yer almayan uretim malzemeleri Malzemeler sayfasindaki tek kaynagi kullanir
        for kod, m in self.malzemeler.items():
            if m.tip == "URETIM" and not m.alternatifler and m.kaynak:
                m.alternatifler = [Alternatif(m.kaynak, m.sn, 0)]
            if m.alternatifler and not m.kaynak:
                m.kaynak = m.alternatifler[0].kaynak

        # ---- EKIP: bir isi ayni anda kac kisinin yapmasi gerektigi ----
        # Kisi tipli kaynaklarda gerekli kisi kadar kisi AYNI GUN AYNI ISE
        # birlikte rezerve edilir. Grupta o kadar kisi yoksa ekip kuculur ve
        # sure oranla uzar; toplam adam-saat her durumda korunur.
        for kod, m in self.malzemeler.items():
            if m.gerekli_kisi <= 1:
                continue
            kisiler = [a for a in m.alternatifler
                       if a.kaynak in self.kaynaklar
                       and self.kaynaklar[a.kaynak].tip == "KISI"]
            if not kisiler:
                # kalip / makine / firin: operator ihtiyaci ek kaynak katsayisiyla
                # modellenir, ayrica ekip kurulmaz
                continue
            m.ekip = min(m.gerekli_kisi, len(kisiler))
            m.ekip_carpan = m.gerekli_kisi / float(m.ekip)
            if m.ekip < m.gerekli_kisi:
                self.uyarilar.append(
                    "'%s' isi %d kisi istiyor ama '%s' grubunda %d kisi tanimli - "
                    "ekip %d kisiye dusuruldu, sure %.2f kat uzatildi "
                    "(adam-saat korunur)."
                    % (kod, m.gerekli_kisi, m.grup or "?", len(kisiler),
                       m.ekip, m.ekip_carpan))

        # --- urun agaci: (a) ust-alt cifti, (b) seviyeli girintili liste ---

        baglar = []          # (ust, alt, oran, kaynak_sayfa)
        ham_agac = sayfa_oku(wb, "UrunAgaci", esl)
        for satir in ham_agac:
            ust = _s(al(satir, "ust_malzeme"))
            alt = _s(al(satir, "alt_malzeme"))
            if not ust or not alt:
                continue
            oran = _f(al(satir, "oran", al(satir, "miktar")), 1.0) or 1.0
            baglar.append((ust, alt, oran * (1.0 + _f(al(satir, "fire"), 0.0) / 100.0),
                           "UrunAgaci"))
        self._baslik_kontrol("UrunAgaci", ham_agac, baglar, "Ust Malzeme / Alt Malzeme")
        baglar += self._seviyeli_agac_oku(wb, esl)

        gorulen = {}
        for ust, alt, oran, nereden in baglar:
            for k in (ust, alt):
                if k not in self.malzemeler:
                    raise ValueError("Urun agacinda (%s) tanimsiz malzeme: '%s'. "
                                     "Once Malzemeler sayfasina ekle." % (nereden, k))
            if (ust, alt) in gorulen:
                if abs(gorulen[(ust, alt)] - oran) > 1e-9:
                    raise ValueError("Urun agacinda celiskili oran: %s -> %s "
                                     "(%g ve %g)" % (ust, alt, gorulen[(ust, alt)], oran))
                continue                       # ayni bag iki formatta da yazilmis, sorun yok
            gorulen[(ust, alt)] = oran
            self.malzemeler[ust].cocuklar.append((alt, oran))
            self.malzemeler[alt].ebeveynler.append(ust)

        # --- talepler ---
        self.talepler = []
        ham_talep = sayfa_oku(wb, "Talepler", esl)
        for satir in ham_talep:
            m = _s(al(satir, "malzeme_kodu"))
            if not m:
                continue
            self.talepler.append({
                "belge": _s(al(satir, "belge")),
                "kalem": _s(al(satir, "kalem")),
                "malzeme": m,
                "miktar": _f(al(satir, "miktar")),
                "birim": _s(al(satir, "birim")),
                "tarih": _d(al(satir, "tarih")),
                "oncelik": _f(al(satir, "oncelik"), 0),
            })

        # --- acik siparisler ---
        self.acik = []
        ham_acik = [r for r in sayfa_oku(wb, "AcikSiparisler", esl)
                    if not _s(al(r, "belge")).upper().startswith("ORNEK")]
        for satir in ham_acik:
            belge = _s(al(satir, "belge"))
            m = _s(al(satir, "malzeme_kodu"))
            if not m or belge.upper().startswith("ORNEK"):
                continue
            self.acik.append({
                "belge": belge, "malzeme": m,
                "miktar": _f(al(satir, "miktar")),
                "tarih": _d(al(satir, "tarih")),
                "aciklama": _s(al(satir, "aciklama")),
            })

        self._baslik_kontrol("AcikSiparisler", ham_acik, self.acik, "Malzeme Kodu")
        self._baslik_kontrol("Talepler", ham_talep, self.talepler, "Malzeme Kodu")
        self._llc_hesapla()
        self._dogrula()

    # ------------------------------------------------------------------ #
    @staticmethod
    def _baslik_kontrol(sayfa, ham, cozulen, beklenen_alan):
        """Dolu satir var ama hicbiri cozulemediyse -> baslik taninmamis demektir."""
        if ham and not cozulen:
            raise ValueError(
                "'%s' sayfasinda %d dolu satir var ama hicbiri okunamadi: "
                "'%s' kolonu bulunamadi.\n"
                "  Okunan basliklar: %s\n"
                "  Cozum: basligi duzelt ya da KolonEslestirme sayfasina ekle "
                "(Sayfa=%s | Dosyadaki Baslik=<senin basligin> | Program Alani=%s)."
                % (sayfa, len(ham), beklenen_alan,
                   ", ".join(OKUNAN_BASLIKLAR.get(sayfa, [])) or "(yok)",
                   sayfa, beklenen_alan))

    # ------------------------------------------------------------------ #
    def _seviyeli_agac_oku(self, wb, esl):
        """
        'UrunAgaciSeviyeli' sayfasi: Seviye | Malzeme | Miktar | Fire %
        Seviye 0 = nihai urun, 1 = onun alti, 2 = onun alti ...
        Bir satirin ebeveyni, ustundeki en yakin (seviye-1) satiridir.
        """
        satirlar = sayfa_oku(wb, "UrunAgaciSeviyeli", esl)
        if not satirlar:
            return []
        baglar = []
        yol = {}                              # seviye -> o seviyedeki son malzeme
        onceki = None
        for n, satir in enumerate(satirlar, start=1):
            kod = _s(al(satir, "malzeme_kodu"))
            if not kod:
                continue
            ham = al(satir, "seviye")
            if ham is None or _s(ham) == "":
                raise ValueError("UrunAgaciSeviyeli %d. satir: Seviye bos ('%s')." % (n, kod))
            sev = int(_f(ham, 0))
            if sev < 0:
                raise ValueError("UrunAgaciSeviyeli %d. satir: Seviye negatif olamaz." % n)
            if onceki is None and sev != 0:
                raise ValueError("UrunAgaciSeviyeli: ilk satir Seviye 0 (nihai urun) olmali, "
                                 "'%s' seviye %d yazilmis." % (kod, sev))
            if onceki is not None and sev > onceki + 1:
                raise ValueError("UrunAgaciSeviyeli %d. satir ('%s'): seviye %d'den %d'e atladi. "
                                 "Seviye birer birer artmali." % (n, kod, onceki, sev))
            if sev > 0:
                ust = yol.get(sev - 1)
                if not ust:
                    raise ValueError("UrunAgaciSeviyeli %d. satir ('%s'): ust seviye "
                                     "bulunamadi." % (n, kod))
                oran = _f(al(satir, "oran", al(satir, "miktar")), 1.0) or 1.0
                baglar.append((ust, kod, oran * (1.0 + _f(al(satir, "fire"), 0.0) / 100.0),
                               "UrunAgaciSeviyeli"))
            yol[sev] = kod
            for daha_derin in [s for s in list(yol) if s > sev]:
                del yol[daha_derin]
            onceki = sev
        return baglar

    # ------------------------------------------------------------------ #
    def _llc_hesapla(self):
        """Low level code: bir malzemenin agactaki en derin seviyesi."""
        def derinlik(kod, iz):
            if kod in iz:
                raise ValueError("Urun agacinda dongu: %s" % " -> ".join(list(iz) + [kod]))
            m = self.malzemeler[kod]
            if not m.cocuklar:
                return 0
            return 1 + max(derinlik(c, iz | {kod}) for c, _ in m.cocuklar)

        kokler = [k for k, m in self.malzemeler.items() if not m.ebeveynler]
        if not kokler and any(m.cocuklar for m in self.malzemeler.values()):
            raise ValueError("Urun agacinda dongu var: hicbir malzeme en ust seviyede degil "
                             "(her malzemenin bir ebeveyni var).")
        for k in self.malzemeler:
            derinlik(k, frozenset())

        def ata(kod, sev):
            m = self.malzemeler[kod]
            if sev > m.llc:
                m.llc = sev
            for c, _ in m.cocuklar:
                ata(c, sev + 1)
        for k in kokler:
            ata(k, 0)

    def _dogrula(self):
        h = []
        for kod, m in self.malzemeler.items():
            if m.tip == "URETIM":
                if not m.alternatifler:
                    h.append("%s: URETIM tipinde ama hicbir kaynak tanimli degil "
                             "(Malzemeler>Kaynak veya Rotalar sayfasi)" % kod)
                for a in m.alternatifler:
                    if a.kaynak not in self.kaynaklar:
                        h.append("%s: '%s' kaynagi Kaynaklar sayfasinda yok" % (kod, a.kaynak))
                    if a.sn <= 0:
                        h.append("%s / %s: Saniye/Adet 0 veya bos" % (kod, a.kaynak))
                    kk = self.kaynaklar.get(a.kaynak)
                    if kk and kk.tip == "KALIP" and not kk.makineler:
                        h.append("%s: '%s' KALIP tipinde ama KalipMakine sayfasinda "
                                 "hicbir makineye bagli degil" % (kod, a.kaynak))
        for kod, kk in self.kaynaklar.items():
            if kk.ek_kaynak:
                if kk.ek_kaynak not in self.kaynaklar:
                    h.append("Kaynaklar '%s': Ek Kaynak '%s' tanimli degil"
                             % (kod, kk.ek_kaynak))
                elif self.kaynaklar[kk.ek_kaynak].ek_kaynak:
                    h.append("Kaynaklar '%s': ek kaynak zinciri destekleniyor degil "
                             "('%s' de bir ek kaynak gosteriyor)" % (kod, kk.ek_kaynak))
                elif kk.ek_kaynak == kod:
                    h.append("Kaynaklar '%s': kendini ek kaynak gosteriyor" % kod)
        for t in self.talepler:
            etiket = "Talep %s/%s" % (t["belge"] or "?", t["kalem"] or "?")
            if t["malzeme"] not in self.malzemeler:
                h.append("%s: '%s' Malzemeler sayfasinda tanimli degil" % (etiket, t["malzeme"]))
            if t["tarih"] is None:
                h.append("%s: teslim tarihi okunamadi (Teslim Tarihi kolonu)" % etiket)
            if t["miktar"] <= 0:
                h.append("%s: miktar 0 veya okunamadi (Miktar kolonu)" % etiket)
        for a in self.acik:
            if a["malzeme"] not in self.malzemeler:
                h.append("AcikSiparisler '%s': '%s' Malzemeler sayfasinda tanimli degil"
                         % (a["belge"], a["malzeme"]))
        if not self.talepler:
            h.append("Talepler sayfasi bos - planlanacak siparis yok.")
        if not self.malzemeler:
            h.append("Malzemeler sayfasi bos.")
        koksuz = [k for k, m in self.malzemeler.items()
                  if not m.ebeveynler and not m.cocuklar
                  and k not in {t["malzeme"] for t in self.talepler}]
        for k in koksuz:
            self.uyarilar.append("Malzeme '%s' hicbir urun agacinda ve hicbir siparişte yok - "
                                 "plana girmeyecek." % k)
        if h:
            raise ValueError("Girdi hatalari:\n  - " + "\n  - ".join(h))

    # ------------------------------------------------------------------ #
    def gun_listesi(self, n=None):
        n = n or self.ufuk
        return [self.baslangic + dt.timedelta(days=i) for i in range(n)]

    def aktarma_gecikmesi(self):
        return 0 if self.aktarma_kurali == "ADET" else 1

    # ================================================================== #
    #  1) MRP PATLATMA + NETLEME
    # ================================================================== #
    def patlat(self):
        """Cok seviyeli patlatma. Her malzeme icin planli emir (Is) listesi uretir."""
        brut = defaultdict(list)             # malzeme -> [(tarih, miktar, etiket, oncelik)]
        for t in self.talepler:
            brut[t["malzeme"]].append(
                (t["tarih"], t["miktar"], "%s/%s" % (t["belge"], t["kalem"]), t["oncelik"]))

        self.isler = defaultdict(list)
        self.satinalma = defaultdict(list)   # satin alinan malzemeler icin net ihtiyac
        self.baslangic_stok = {k: m.stok for k, m in self.malzemeler.items()}

        # acik siparisleri (giris) stoga tarih sirasiyla ekleyecegiz
        girisler = defaultdict(list)
        for a in self.acik:
            girisler[a["malzeme"]].append((a["tarih"] or self.baslangic, a["miktar"], a["belge"]))

        for kod in sorted(self.malzemeler, key=lambda k: self.malzemeler[k].llc):
            m = self.malzemeler[kod]
            ihtiyaclar = sorted(brut[kod], key=lambda x: (x[3], x[0], str(x[2])))
            if not ihtiyaclar:
                continue

            stok = m.stok - m.emniyet
            gir = sorted(girisler[kod], key=lambda x: x[0])
            gi = 0
            for tarih, miktar, etiket, oncelik in ihtiyaclar:
                while gi < len(gir) and gir[gi][0] <= tarih:
                    stok += gir[gi][1]
                    gi += 1
                if stok >= miktar - EPS:
                    stok -= miktar
                    continue
                net = miktar - max(stok, 0.0)
                stok = min(stok, 0.0)
                if m.tip == "SATINALMA":
                    self.satinalma[kod].append({"tarih": tarih, "miktar": net, "etiket": etiket})
                    continue
                i = Is(kod, net, tarih, etiket, m.llc, oncelik)
                self.isler[kod].append(i)
                for c, oran in m.cocuklar:
                    brut[c].append((tarih, net * oran, etiket, oncelik))

        for kod in self.isler:
            self.isler[kod].sort(key=lambda i: (i.ihtiyac_tarihi, str(i.kaynak_talep)))
        return self.isler

    # ================================================================== #
    #  2) GERI PLANLAMA  (teslim tarihinden geriye, sonlu kapasite)
    # ================================================================== #
    def geri_planla(self):
        """
        Ust seviye hangi gun ne kadar uretiyorsa, alt seviye o miktari
        (aktarma gecikmesi kadar) once hazir etmek zorunda. Boylece seviyeler
        ust uste biner.
        """
        gecikme = self.aktarma_gecikmesi()
        kalan_kap = {}                        # (kaynak, gun) -> kalan saniye
        self.geri_uretim = defaultdict(lambda: defaultdict(float))   # malzeme -> gun -> miktar
        self.geri_bilgi = {}                  # is.no -> (bas, bit)

        # malzeme -> [(bitis_tarihi, miktar, is)]  seklinde talep kuyrugu
        talep = defaultdict(list)
        for kod, isler in self.isler.items():
            for i in isler:
                talep[kod].append([i.ihtiyac_tarihi, i.miktar, i])

        def gun_havuzu(g):
            """O gunun kalan kapasitesi - tum kaynaklar icin, gunler arasi saklanir."""
            if g not in kalan_kap:
                kalan_kap[g] = {k: kk.kapasite_sn(g) for k, kk in self.kaynaklar.items()}
            return kalan_kap[g]

        geri_atama = {}
        alt_talep = defaultdict(list)
        for kod in sorted(self.malzemeler, key=lambda k: self.malzemeler[k].llc):
            m = self.malzemeler[kod]
            if m.tip != "URETIM":
                continue
            kuyruk = talep[kod] + alt_talep[kod]
            if not kuyruk:
                continue
            # en gec teslimden basla; ayni tarihte oncelikli is once yerlesir
            kuyruk.sort(key=lambda x: (x[0], -getattr(x[2], "oncelik", 0.0)), reverse=True)
            imlec = None
            for bitis, miktar, sahip in kuyruk:
                g = bitis if imlec is None else min(imlec, bitis)
                kalan = miktar
                son_gun, ilk_gun = None, None
                guvenlik = 0
                while kalan > EPS:
                    guvenlik += 1
                    if guvenlik > 20000:
                        break
                    havuz = gun_havuzu(g)
                    if any(havuz.get(a.kaynak, 0.0) > EPS for a in m.alternatifler):
                        paylar = self._dagit(m, kalan, havuz, g, None, geri_atama)
                        adet = sum(p[2] for p in paylar)
                        if adet > EPS:
                            self.geri_uretim[kod][g] += adet
                            kalan -= adet
                            if son_gun is None:
                                son_gun = g
                            ilk_gun = g
                            # alt seviye bu miktari 'gecikme' gun once hazirlamali
                            for cc, oran in m.cocuklar:
                                if self.malzemeler[cc].tip == "URETIM":
                                    alt_talep[cc].append(
                                        [g - dt.timedelta(days=gecikme), adet * oran, sahip])
                    g -= dt.timedelta(days=1)
                imlec = g + dt.timedelta(days=1)
                if isinstance(sahip, Is) and son_gun is not None:
                    onceki = self.geri_bilgi.get(sahip.no)
                    bas = ilk_gun if onceki is None else min(onceki[0], ilk_gun)
                    bit = son_gun if onceki is None else max(onceki[1], son_gun)
                    self.geri_bilgi[sahip.no] = (bas, bit)
        return self.geri_uretim

    # ================================================================== #
    #  KAPASITE DAGITIMI: alternatif kaynaklar + kalip -> makine
    # ================================================================== #
    def _makine_sirasi(self, kaynak, gun, atama=None):
        """
        Bir kalibin o gun denenecek makineleri, oncelik sirasiyla.
        YAPISKAN ATAMA: kalip halihazirda bir makineye bagliysa once o denenir;
        ancak o makine yer vermezse (ve minimum kalis suresi dolduysa) siradaki
        makineye tasinir.
        """
        if not kaynak.makineler:
            return [(None, 1.0)]
        sirali = list(kaynak.makineler)
        mevcut = (self.kalip_atama if atama is None else atama).get(kaynak.kod)
        if not mevcut:
            return sirali
        makine, baglandi = mevcut
        oncelikli = [(m, c) for m, c in sirali if m == makine]
        digerleri = [(m, c) for m, c in sirali if m != makine]
        if self.min_kalip_kalis > 0 and (gun - baglandi).days < self.min_kalip_kalis:
            return oncelikli or sirali          # sure dolmadan tasima yok
        return oncelikli + digerleri

    def _gereken_kaynak(self, m, kalan, gun, son_gun):
        """
        Isi teslim tarihine (son_gun) yetistirmek icin ayni anda kac kaynak
        (kisi / ekip / kalip) gerektigini soyler. 1 kisi yetiyorsa 1 doner;
        yetmiyorsa gereken kadar TAKVIYE sayisi doner. 0 = sinir koyma.
        """
        if not m.alternatifler or kalan <= EPS:
            return 1
        alt = m.alternatifler[0]
        kk = self.kaynaklar.get(alt.kaynak)
        sn = alt.sn * (m.ekip_carpan if m.ekip > 1 else 1.0)
        if kk is None or sn <= 0:
            return 1
        if son_gun is None or son_gun <= gun:
            return 0                                   # tarih geldi/gecti: serbest
        # teslime kadar kaynagin kac is gunu kaldi
        is_gunu, g = 0, gun
        while g <= son_gun and is_gunu < 500:
            if kk.calisir(g):
                is_gunu += 1
            g += dt.timedelta(days=1)
        is_gunu = max(1, is_gunu - self.takviye_payi)
        if is_gunu <= 0:
            return 0
        gunluk_adet = kk.kapasite_sn(gun) / sn
        if gunluk_adet <= 0:
            return 0
        gerek = int(math.ceil(kalan / (is_gunu * gunluk_adet) - 1e-9))
        return max(1, min(gerek, len(m.alternatifler)))

    def _dagit(self, m, hedef, havuz, gun, yuk=None, atama=None, gecmis=None,
               gun_isleri=None, sinir=0, kaynak_sinir=0, is_kalan=None, acil=False):
        """
        'hedef' adedi, malzemenin alternatif kaynaklari arasinda oncelik sirasiyla
        dagitir. 1. oncelikli kaynagin o gunku kapasitesi dolunca KALAN MIKTAR
        ayni gun 2. oncelikliye tasar.
        Kaynak bir KALIP ise ayni anda bir MAKINE kapasitesi de tuketilir.
        Dondurur: [(kaynak_kodu, makine_kodu|None, adet, saniye/adet, ek_kaynaklar,
                    ekip), ...]
        ek_kaynaklar: {kaynak: katsayi} - ayni anda tuketilen havuzlar (orn. operator)
        ekip        : [kaynak_kodu] - isi ayni anda yapan DIGER kisiler (bos olabilir)
        """
        if m.ekip > 1:
            return self._dagit_ekip(m, hedef, havuz, gun, yuk, gun_isleri, sinir,
                                    kaynak_sinir)
        paylar = []
        kalan = hedef
        # Is basina kaynak siniri: bir urun bir hafta icinde en fazla kac farkli
        # kisiye/kalibda bolunebilir. Amac isi erken bitirmek icin kalabaliga
        # dagitmak yerine ayni kisiye birkac gun vermek. Teslim tarihi gelmis
        # isler icin faz 1 bu siniri kaldirir (kaynak_sinir=0 gelir).
        hkume = None
        if kaynak_sinir > 0 and yuk is not None:
            hkume = self.hafta_kaynak[(gun.isocalendar()[:2], m.kod)]
        for alt in m.alternatifler:
            if kalan <= EPS:
                break
            kaynak = self.kaynaklar.get(alt.kaynak)
            if kaynak is None:
                continue
            if hkume is not None and alt.kaynak not in hkume and len(hkume) >= kaynak_sinir:
                continue
            # gun ici is degisimi siniri: bu kaynak bugun baska islere basladiysa
            # ve kota dolduysa yeni bir urune gecme, siradaki alternatifi dene
            if sinir > 0 and gun_isleri is not None:
                yapilan = gun_isleri.get(alt.kaynak)
                if yapilan and m.kod not in yapilan and len(yapilan) >= sinir:
                    continue
            # ---- PARTILI KAYNAK (firin): Q adetlik parti T saniyede pisirilir ----
            # Bir parti dolu da olsa yarim da olsa ayni sureyi harcar. Ayni gun ayni
            # firinda ayni urunun acik partisinde yer varsa ek sure harcanmaz.
            # Yarim parti ancak isin kalaninin tamami hazirsa ya da teslim
            # yaklastiysa (acil) baslatilir; yoksa malzemenin birikmesi beklenir.
            if m.parti and yuk is not None and not kaynak.makineler:
                q_parti, t_parti = m.parti
                anah = (gun, alt.kaynak, m.kod)
                bosluk = self.parti_bosluk.get(anah, 0.0)
                q_bos = min(kalan, bosluk)
                kalan2 = kalan - q_bos
                n_tam = int(kalan2 // q_parti + 1e-9)
                artik = kalan2 - n_tam * q_parti
                son_parti = is_kalan is None or kalan2 >= (is_kalan - q_bos) - EPS
                n_ist = n_tam + (1 if artik > EPS and (son_parti or acil) else 0)
                n_sig = int(havuz.get(alt.kaynak, 0.0) // t_parti + 1e-9)
                n = min(n_ist, n_sig)
                q_yeni = min(kalan2, n * q_parti)
                tavan = q_bos + q_yeni
                if tavan <= EPS:
                    continue
                sure = n * t_parti
                self.parti_bosluk[anah] = bosluk - q_bos + (n * q_parti - q_yeni)
                havuz[alt.kaynak] = havuz.get(alt.kaynak, 0.0) - sure
                yuk[alt.kaynak][gun] += sure
                self.parti_sayisi += n
                if gun_isleri is not None:
                    gun_isleri.setdefault(alt.kaynak, set()).add(m.kod)
                paylar.append((alt.kaynak, None, tavan, sure / tavan, {}, []))
                kalan -= tavan
                if hkume is not None:
                    hkume.add(alt.kaynak)
                continue
            for makine_kod, carpan in self._makine_sirasi(kaynak, gun, atama):
                if kalan <= EPS:
                    break
                kap_kaynak = havuz.get(alt.kaynak, 0.0)
                if kap_kaynak <= EPS:
                    break                                   # kalip/kisi doldu
                sn_eff = alt.sn * (carpan if makine_kod else 1.0)
                if sn_eff <= 0:
                    continue
                tavan = min(kalan, kap_kaynak / sn_eff)
                if makine_kod:
                    tavan = min(tavan, havuz.get(makine_kod, 0.0) / sn_eff)
                # ayni anda tuketilen ek kaynaklar (orn. operator havuzu)
                ekler = {}
                for kk in (kaynak, self.kaynaklar.get(makine_kod) if makine_kod else None):
                    if kk is not None and kk.ek_kaynak:
                        ekler[kk.ek_kaynak] = ekler.get(kk.ek_kaynak, 0.0) + kk.ek_katsayi
                # kisi tipli olmayan kaynakta "gerekli kisi" operator ihtiyacini
                # katlar (orn. bir makineye 2 operator)
                if m.gerekli_kisi > 1 and ekler:
                    ekler = {k: v * m.gerekli_kisi for k, v in ekler.items()}
                for ek, kats in ekler.items():
                    tavan = min(tavan, havuz.get(ek, 0.0) / (sn_eff * kats))
                if tavan <= EPS:
                    continue
                havuz[alt.kaynak] = kap_kaynak - tavan * sn_eff
                if yuk is not None:
                    yuk[alt.kaynak][gun] += tavan * sn_eff
                for ek, kats in ekler.items():
                    havuz[ek] = havuz.get(ek, 0.0) - tavan * sn_eff * kats
                    if yuk is not None:
                        yuk[ek][gun] += tavan * sn_eff * kats
                if makine_kod:
                    havuz[makine_kod] = havuz.get(makine_kod, 0.0) - tavan * sn_eff
                    if yuk is not None:
                        yuk[makine_kod][gun] += tavan * sn_eff
                    hedef_atama = self.kalip_atama if atama is None else atama
                    onceki = hedef_atama.get(alt.kaynak)
                    if not onceki or onceki[0] != makine_kod:
                        hedef_atama[alt.kaynak] = (makine_kod, gun)
                        if gecmis is not None:
                            gecmis.append((gun, alt.kaynak, makine_kod,
                                           onceki[0] if onceki else None))
                if gun_isleri is not None:
                    gun_isleri.setdefault(alt.kaynak, set()).add(m.kod)
                paylar.append((alt.kaynak, makine_kod, tavan, sn_eff, dict(ekler), []))
                kalan -= tavan
                if hkume is not None:
                    hkume.add(alt.kaynak)
        return paylar

    def _dagit_ekip(self, m, hedef, havuz, gun, yuk=None, gun_isleri=None, sinir=0,
                    kaynak_sinir=0):
        """
        Bir isi AYNI ANDA m.ekip kisinin yaptigi durum (ERP: CALISAN PERSONEL
        SAYISI > 1). Alternatif listesi oncelik sirasinda taranir, kapasitesi
        olan ilk m.ekip kisi bir EKIP olusturur; ekibin her uyesi ayni sureyi
        harcar. Ekip dolunca kalan miktar icin siradaki kisilerden yeni bir
        ekip kurulur (ayni gun ikinci ekip).
        """
        paylar = []
        kalan = hedef
        aday = [a for a in m.alternatifler
                if a.kaynak in self.kaynaklar
                and self.kaynaklar[a.kaynak].tip == "KISI"]
        i = 0
        while kalan > EPS and i < len(aday):
            ekip = []
            while i < len(aday) and len(ekip) < m.ekip:
                alt = aday[i]
                i += 1
                if havuz.get(alt.kaynak, 0.0) <= EPS:
                    continue
                if sinir > 0 and gun_isleri is not None:
                    yapilan = gun_isleri.get(alt.kaynak)
                    if yapilan and m.kod not in yapilan and len(yapilan) >= sinir:
                        continue
                ekip.append(alt)
            if len(ekip) < m.ekip:
                break                                   # tam ekip kurulamadi
            sn_eff = ekip[0].sn * m.ekip_carpan
            if sn_eff <= 0:
                break
            tavan = kalan
            for alt in ekip:
                tavan = min(tavan, havuz.get(alt.kaynak, 0.0) / sn_eff)
            if tavan <= EPS:
                continue
            for alt in ekip:
                havuz[alt.kaynak] = havuz.get(alt.kaynak, 0.0) - tavan * sn_eff
                if yuk is not None:
                    yuk[alt.kaynak][gun] += tavan * sn_eff
                if gun_isleri is not None:
                    gun_isleri.setdefault(alt.kaynak, set()).add(m.kod)
            paylar.append((ekip[0].kaynak, None, tavan, sn_eff, {},
                           [a.kaynak for a in ekip[1:]]))
            kalan -= tavan
            if kaynak_sinir > 0 and len(paylar) >= kaynak_sinir:
                break                                   # hafta icinde tek ekip yetsin
        return paylar

    # ================================================================== #
    #  3) ILERI PLANLAMA  (bugunden ileri, gun gun simulasyon)
    # ================================================================== #
    def ileri_planla(self):
        gecikme = self.aktarma_gecikmesi()
        stok = {k: m.stok for k, m in self.malzemeler.items()}
        gelen = defaultdict(lambda: defaultdict(float))     # gun -> malzeme -> miktar
        acik_hareket = []
        for a in self.acik:
            g = a["tarih"] or self.baslangic
            if g < self.baslangic:
                g = self.baslangic                          # gecmis tarihli acik siparis bugune cekilir
            gelen[g][a["malzeme"]] += a["miktar"]
            acik_hareket.append((a["malzeme"], g, a["belge"], a["miktar"]))

        self.uretim = defaultdict(lambda: defaultdict(float))    # malzeme -> gun -> miktar
        self.tuketim = defaultdict(lambda: defaultdict(float))   # malzeme -> gun -> miktar
        self.yuk = defaultdict(lambda: defaultdict(float))       # kaynak -> gun -> saniye
        self.teslim = []                                         # (gun, talep, miktar)
        self.hareket = defaultdict(list)                         # malzeme -> [(gun, tip, aciklama, miktar)]
        self.satinalma_ileri = defaultdict(lambda: defaultdict(float))
        self.giris = defaultdict(lambda: defaultdict(float))      # malzeme -> gun -> stoga giren
        self.hafta_kaynak = defaultdict(set)   # (yil-hafta, malzeme) -> kullanilan kaynaklar
        self.parti_bosluk = {}                 # (gun, firin, malzeme) -> acik partide kalan yer
        self.parti_sayisi = 0
        self.hafta_sinir = {}                  # (yil-hafta, malzeme) -> uygulanan kaynak siniri
        self.uretim_detay = []          # (gun, malzeme, kaynak, makine, adet, saat,
                                        #  talep, is_no, ek_kaynak, ekip)
        self.kalip_atama = {}           # kalip -> (makine, baglandigi gun)
        self.kalip_gecmisi = []         # (gun, kalip, yeni_makine, eski_makine)
        self.kalan_kapasite = {}        # gun -> {kaynak: kalan saniye}   (kontrol icin)
        self.gun_is_kaydi = {}          # gun -> {kaynak: {o gun calisilan urunler}}
        for kod, g, belge, q in acik_hareket:
            self.hareket[kod].append(
                (g, "Acik Siparis", belge, q))
            self.giris[kod][g] += q

        # uretim sirasi: en derin seviye once (alt seviye once uretir)
        uretim_sirasi = sorted(
            [k for k, m in self.malzemeler.items() if m.tip == "URETIM"],
            key=lambda k: -self.malzemeler[k].llc)

        # talepler
        acik_talepler = sorted(
            [dict(t, kalan=t["miktar"], karsilanan_tarih=None) for t in self.talepler],
            key=lambda t: (t["oncelik"], t["tarih"], t["belge"], t["kalem"]))

        # Sevk sirasi. AYNI KAYNAGI PAYLASAN farkli malzemeler ayni kapasite
        # havuzundan yer alir; isler global olarak siralanip yarisir.
        #   - once oncelik (kucuk = onde), sonra ihtiyac tarihi (EDD),
        #   - esitlikte derin seviye once (alt seviye beslenmeli),
        #   - sonra is numarasi (kararli siralama)
        def anahtar(i):
            if self.siralama == "FIFO":
                return (i.oncelik, i.no)
            return (i.oncelik, i.ihtiyac_tarihi, -self.malzemeler[i.malzeme].llc, i.no)

        tum_isler = [i for kod in uretim_sirasi for i in self.isler.get(kod, [])]
        tum_isler.sort(key=anahtar)
        en_derin = max([m.llc for m in self.malzemeler.values()] or [0])

        bitis_gun = None          # son URETIM gunu
        sevk_gun = None           # son SEVKIYAT gunu
        for gun in self.gun_listesi():
            # a) girisler
            for m, q in gelen.pop(gun, {}).items():
                stok[m] = stok.get(m, 0.0) + q

            # b) o gunun kapasite havuzu  (kaynak -> kalan saniye)
            havuz = {k: kaynak.kapasite_sn(gun) for k, kaynak in self.kaynaklar.items()}
            gun_isleri = {}
            kalan_is = any(i.kalan > EPS for isler in self.isler.values() for i in isler)
            if kalan_is and any(v > EPS for v in havuz.values()):
                # Aktarma kurali ADET ise ayni gun icinde seviyeler arasi akis olabilir;
                # bu yuzden ilerleme durana kadar tekrar gecilir. GUNLUK/SABIT kuralinda
                # ikinci gecis zaten bir sey uretmez (malzeme ertesi gune acilir).
                for _ in range(en_derin + 1 if gecikme == 0 else 1):
                    ilerleme = False
                    # faz 0: gun ici is degisimi sinirina uy
                    # faz 1: sadece teslim tarihi gelmis/gecmis isler icin siniri kaldir
                    fazlar = (0, 1) if (self.gunluk_is_cesidi > 0
                                        or self.is_basina_kaynak > 0) else (0,)
                    for faz in fazlar:
                      sinir = self.gunluk_is_cesidi if faz == 0 else 0
                      kaynak_sinir = self.is_basina_kaynak if faz == 0 else 0
                      for i in tum_isler:
                        if i.kalan <= EPS:
                            continue
                        if faz == 1 and i.ihtiyac_tarihi > gun:
                            continue
                        m = self.malzemeler[i.malzeme]
                        # TAKVIYE: is tek kisiyle teslime yetismiyorsa sinir
                        # gereken kadar genisletilir ("bolme ama gecikme de olmasin")
                        is_sinir = kaynak_sinir
                        if is_sinir > 0:
                            gerek = self._gereken_kaynak(m, i.kalan, gun,
                                                         i.ihtiyac_tarihi)
                            is_sinir = 0 if gerek == 0 else max(is_sinir, gerek)
                        # o hafta/malzeme icin fiilen gecerli olan sinir
                        hk = (gun.isocalendar()[:2], i.malzeme)
                        onc = self.hafta_sinir.get(hk)
                        self.hafta_sinir[hk] = (0 if (is_sinir == 0 or onc == 0)
                                                else max(onc or 0, is_sinir))
                        # malzeme (bilesen) siniri
                        tavan = i.kalan
                        for c, oran in m.cocuklar:
                            if self.malzemeler[c].tip == "SATINALMA" and not self.satinalma_kisit:
                                continue
                            tavan = min(tavan, stok.get(c, 0.0) / oran if oran > 0 else tavan)
                        if tavan <= EPS:
                            continue
                        # kapasiteyi alternatifler (ve kalip->makine) arasinda dagit
                        paylar = self._dagit(m, tavan, havuz, gun, self.yuk,
                                             self.kalip_atama, self.kalip_gecmisi,
                                             gun_isleri, sinir, is_sinir,
                                             is_kalan=i.kalan,
                                             acil=(i.ihtiyac_tarihi is None or
                                                   (i.ihtiyac_tarihi - gun).days
                                                   <= self.takviye_payi))
                        adet = sum(p[2] for p in paylar)
                        if adet <= EPS:
                            continue
                        for c, oran in m.cocuklar:
                            cm = self.malzemeler[c]
                            gerek = adet * oran
                            eldeki = min(stok.get(c, 0.0), gerek)
                            stok[c] = stok.get(c, 0.0) - gerek
                            if cm.tip == "SATINALMA" and not self.satinalma_kisit \
                                    and gerek - eldeki > EPS:
                                self.satinalma_ileri[c][gun] += gerek - eldeki
                            self.tuketim[c][gun] += gerek
                            self.hareket[c].append(
                                (gun, "Ikincil Ihtiyac",
                                 "%s / %s" % (i.malzeme, i.kaynak_talep), -gerek))
                        for kaynak_kod, makine_kod, pay, sn_eff, ekler, ekip in paylar:
                            self.uretim_detay.append(
                                (gun, i.malzeme, kaynak_kod, makine_kod, pay,
                                 pay * sn_eff / 3600.0, i.kaynak_talep, i.no,
                                 ", ".join(sorted(ekler)) if ekler else "",
                                 ", ".join(ekip)))
                        i.kalan -= adet
                        i.gunluk[gun] += adet
                        if i.baslangic is None:
                            i.baslangic = gun
                        i.bitis = gun
                        hazir = gun + dt.timedelta(days=gecikme)
                        if gecikme == 0:
                            stok[i.malzeme] = stok.get(i.malzeme, 0.0) + adet
                        else:
                            gelen[hazir][i.malzeme] += adet
                        self.uretim[i.malzeme][gun] += adet
                        self.giris[i.malzeme][hazir] += adet
                        self.hareket[i.malzeme].append(
                            (hazir, "Uretim", "%s (is %d)" % (i.kaynak_talep, i.no), adet))
                        bitis_gun = gun
                        ilerleme = True
                    if not ilerleme:
                        break

            self.kalan_kapasite[gun] = dict(havuz)
            self.gun_is_kaydi[gun] = {k: set(v) for k, v in gun_isleri.items()}

            # c) sevkiyat
            for t in acik_talepler:
                if t["kalan"] <= EPS:
                    continue
                # Teslim tarihinden once sevkiyat yok: mal erken biterse stokta bekler.
                if not self.erken_sevkiyat and t["tarih"] and gun < t["tarih"]:
                    continue
                mevcut = stok.get(t["malzeme"], 0.0)
                if mevcut <= EPS:
                    continue
                q = min(t["kalan"], mevcut)
                stok[t["malzeme"]] = mevcut - q
                t["kalan"] -= q
                self.teslim.append((gun, t, q))
                self.hareket[t["malzeme"]].append(
                    (gun, "Satis Siparisi", "%s/%s" % (t["belge"], t["kalem"]), -q))
                if t["kalan"] <= EPS:
                    t["karsilanan_tarih"] = gun

            if all(t["kalan"] <= EPS for t in acik_talepler) and \
               all(i.kalan <= EPS for isler in self.isler.values() for i in isler):
                sevk_gun = gun
                break

        self.son_gun = bitis_gun or self.baslangic
        self.son_sevk = sevk_gun or self.son_gun
        self.talep_sonuc = acik_talepler
        self.son_stok = stok
        return acik_talepler


# ====================================================================== #
#  EXCEL CIKTISI
# ====================================================================== #

BASLIK_FILL = PatternFill("solid", fgColor="1F3864")
BASLIK_FONT = Font(color="FFFFFF", bold=True, size=10)
INCE = Side(style="thin", color="D9D9D9")
KENAR = Border(left=INCE, right=INCE, top=INCE, bottom=INCE)
IYI = PatternFill("solid", fgColor="C6EFCE")
KOTU = PatternFill("solid", fgColor="FFC7CE")
UYARI = PatternFill("solid", fgColor="FFEB9C")


def tablo(ws, basliklar, satirlar, genislikler=None, formatlar=None, baslik_satiri=1):
    for j, b in enumerate(basliklar, start=1):
        c = ws.cell(row=baslik_satiri, column=j, value=b)
        c.fill = BASLIK_FILL
        c.font = BASLIK_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = KENAR
    for i, satir in enumerate(satirlar, start=baslik_satiri + 1):
        for j, v in enumerate(satir, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.border = KENAR
            if isinstance(v, dt.date) and not isinstance(v, dt.datetime):
                c.number_format = "DD.MM.YYYY"
            elif formatlar and basliklar[j - 1] in formatlar:
                c.number_format = formatlar[basliklar[j - 1]]
    if genislikler:
        for j, g in enumerate(genislikler, start=1):
            ws.column_dimensions[get_column_letter(j)].width = g
    ws.freeze_panes = ws.cell(row=baslik_satiri + 1, column=1)
    ws.auto_filter.ref = "A%d:%s%d" % (baslik_satiri, get_column_letter(len(basliklar)),
                                       baslik_satiri + max(len(satirlar), 1))
    return baslik_satiri + len(satirlar)


def rapor_yaz(mdl, yol):
    wb = Workbook()

    # ---------------- OZET ---------------- #
    ws = wb.active
    ws.title = "Ozet"
    toplam_talep = sum(t["miktar"] for t in mdl.talepler)
    zamaninda = [t for t in mdl.talep_sonuc if t["karsilanan_tarih"] and t["karsilanan_tarih"] <= t["tarih"]]
    gecikenler = [t for t in mdl.talep_sonuc if t["karsilanan_tarih"] and t["karsilanan_tarih"] > t["tarih"]]
    karsilanmayan = [t for t in mdl.talep_sonuc if not t["karsilanan_tarih"]]
    max_gecikme = max([(t["karsilanan_tarih"] - t["tarih"]).days for t in gecikenler] or [0])

    # kaynak yuku
    kaynak_ozet = []
    for kod, kaynak in mdl.kaynaklar.items():
        toplam_sn = sum(mdl.yuk[kod].values())
        gunler = [g for g in mdl.gun_listesi((mdl.son_gun - mdl.baslangic).days + 1)
                  if kaynak.calisir(g)]
        mevcut_sn = sum(kaynak.kapasite_sn(g) for g in gunler)
        calisilan = len([g for g in mdl.yuk[kod] if mdl.yuk[kod][g] > EPS])
        kaynak_ozet.append((kod, toplam_sn / 3600.0, mevcut_sn / 3600.0,
                            (toplam_sn / mevcut_sn * 100.0) if mevcut_sn else 0.0, calisilan,
                            kaynak.paralel,
                            (toplam_sn / 3600.0 / (kaynak.saat_gun * calisilan))
                            if kaynak.saat_gun and calisilan else 0.0))
    darbogaz = max(kaynak_ozet, key=lambda x: x[3]) if kaynak_ozet else None

    satirlar = [
        ["Plan baslangic tarihi", mdl.baslangic, ""],
        ["Aktarma kurali", mdl.aktarma_kurali,
         "Alt seviyenin 1 gunluk uretimi birikince ust seviye baslar" if mdl.aktarma_kurali == "GUNLUK"
         else ("Adet adet akis" if mdl.aktarma_kurali == "ADET" else "Sabit parti: %g" % mdl.aktarma_partisi)],
        ["", "", ""],
        ["Talep kalemi sayisi", len(mdl.talepler), ""],
        ["Toplam talep miktari", toplam_talep, ""],
        ["Zamaninda karsilanan kalem", len(zamaninda), ""],
        ["Geciken kalem", len(gecikenler), "En buyuk gecikme: %d gun" % max_gecikme],
        ["Ufuk icinde karsilanamayan", len(karsilanmayan), ""],
        ["", "", ""],
        ["Son uretim gunu (ileri plan)", mdl.son_gun, ""],
        ["Son sevkiyat gunu", getattr(mdl, "son_sevk", mdl.son_gun),
         "Erken sevkiyat kapaliysa son teslim tarihine kadar uzar"],
        ["Darbogaz kaynak", darbogaz[0] if darbogaz else "-",
         "%.1f%% doluluk" % darbogaz[3] if darbogaz else ""],
        ["", "", ""],
    ]
    son = tablo(ws, ["Gosterge", "Deger", "Not"], satirlar, [30, 20, 60],
                {"Deger": "#,##0.###"})

    ws.cell(row=son + 2, column=1,
            value="KAYNAK YUKU  (en dolu ustte)").font = Font(bold=True, size=11)
    son = tablo(ws, ["Kaynak", "Kaynak / Kisi Adi", "Tip", "Paralel Adet",
                     "Yuklenen (saat)", "Mevcut kapasite (saat)", "Doluluk %",
                     "Calisilan gun", "Esdeger kisi/makine"],
                [[k, mdl.kaynaklar[k].aciklama[:52], mdl.kaynaklar[k].tip, par,
                  round(a, 1), round(b, 1), round(c, 1), d, round(esd, 2)]
                 for k, a, b, c, d, par, esd in sorted(kaynak_ozet, key=lambda x: -x[3])],
                [20, 46, 11, 12, 14, 19, 11, 13, 18],
                {"Yuklenen (saat)": "#,##0.0", "Mevcut kapasite (saat)": "#,##0.0",
                 "Doluluk %": "#,##0.0", "Esdeger kisi/makine": "#,##0.00"},
                baslik_satiri=son + 3)
    ws.cell(row=son + 1, column=1,
            value=("Paralel Adet > 1 olan kaynaklar HAVUZDUR (orn. enjeksiyon operatorleri). "
                   "'Esdeger kisi/makine' = calisilan gunlerde gunde ortalama kac kisi/makine "
                   "mesgul; havuzun kac kisisinin fiilen gerektigini gosterir.")
            ).font = Font(italic=True, size=9, color="808080")

    # urun bazinda ozet (cok urunlu calismalarda asil bakilacak tablo)
    urun_ozet = []
    for kod in sorted({t["malzeme"] for t in mdl.talepler},
                      key=lambda k: mdl.malzemeler[k].llc):
        kalemler = [t for t in mdl.talep_sonuc if t["malzeme"] == kod]
        g = [t for t in kalemler if t["karsilanan_tarih"] and t["karsilanan_tarih"] > t["tarih"]]
        y = [t for t in kalemler if not t["karsilanan_tarih"]]
        son_t = max([t["karsilanan_tarih"] for t in kalemler if t["karsilanan_tarih"]] or [None])
        urun_ozet.append([kod, mdl.malzemeler[kod].aciklama[:50], len(kalemler),
                          sum(t["miktar"] for t in kalemler),
                          len(kalemler) - len(g) - len(y), len(g), len(y),
                          max([(t["karsilanan_tarih"] - t["tarih"]).days for t in g] or [0]),
                          son_t, mdl.malzemeler[kod].kaynak])
    ws.cell(row=son + 2, column=1, value="URUN BAZINDA OZET").font = Font(bold=True, size=11)
    son = tablo(ws, ["Malzeme", "Aciklama", "Kalem", "Toplam Miktar", "Zamaninda",
                     "Geciken", "Karsilanamayan", "En Buyuk Gecikme (gun)", "Son Teslim", "Kaynak"],
                urun_ozet, None, {"Toplam Miktar": "#,##0.###"}, baslik_satiri=son + 3)

    if mdl.uyarilar:
        ws.cell(row=son + 2, column=1, value="UYARILAR").font = Font(bold=True, size=11)
        for i, u in enumerate(mdl.uyarilar, start=son + 3):
            c = ws.cell(row=i, column=1, value=u)
            c.font = Font(size=9, color="BF8F00")
            ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=5)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = None

    # ---------------- SIPARIS OZETI ---------------- #
    ws = wb.create_sheet("SiparisOzeti")
    satirlar = []
    for t in sorted(mdl.talep_sonuc, key=lambda x: (x["tarih"], x["belge"], x["kalem"])):
        gerekli_bas = None
        for kod, isler in mdl.isler.items():
            for i in isler:
                if i.kaynak_talep == "%s/%s" % (t["belge"], t["kalem"]) and i.no in mdl.geri_bilgi:
                    b = mdl.geri_bilgi[i.no][0]
                    gerekli_bas = b if gerekli_bas is None else min(gerekli_bas, b)
        kt = t["karsilanan_tarih"]
        gecikme = (kt - t["tarih"]).days if kt else None
        durum = ("KARSILANAMADI" if kt is None else
                 ("ZAMANINDA" if gecikme <= 0 else "GECIKME %d gun" % gecikme))
        satirlar.append([t["belge"], t["kalem"], t["malzeme"], t["miktar"], t["tarih"],
                         gerekli_bas,
                         "GECMISTE" if (gerekli_bas and gerekli_bas < mdl.baslangic) else "OK",
                         kt, gecikme if gecikme is not None else "", durum])
    son = tablo(ws, ["Belge", "Kalem", "Malzeme", "Miktar", "Teslim Tarihi",
                     "Geri Plan: Gereken Baslangic", "Geri Plan Durumu",
                     "Ileri Plan: Bitis", "Gecikme (gun)", "Durum"],
                satirlar, [16, 8, 24, 12, 14, 22, 16, 16, 13, 20],
                {"Miktar": "#,##0.###", "Gecikme (gun)": "#,##0"})
    ws.conditional_formatting.add("J2:J%d" % son,
                                  CellIsRule(operator="equal", formula=['"ZAMANINDA"'], fill=IYI))
    ws.conditional_formatting.add("I2:I%d" % son,
                                  CellIsRule(operator="greaterThan", formula=["0"], fill=KOTU))
    ws.conditional_formatting.add("G2:G%d" % son,
                                  CellIsRule(operator="equal", formula=['"GECMISTE"'], fill=UYARI))

    # ---------------- URETIM PLANI (ileri) ---------------- #
    # Ayni malzeme ayni gun birden fazla kaynak/makinede calisabildigi icin
    # (kapasite tasmasi) satirlar kaynak+makine kiriliminda birlestirilir.
    ws = wb.create_sheet("UretimPlani")
    birlesik = OrderedDict()
    for gun, kod, kaynak, makine, adet, saat, talep, isno, ek, ekip in mdl.uretim_detay:
        a = (gun, kod, kaynak, makine, ek, ekip)
        if a not in birlesik:
            birlesik[a] = [0.0, 0.0]
        birlesik[a][0] += adet
        birlesik[a][1] += saat
    kum = defaultdict(float)
    satirlar = []
    for (gun, kod, kaynak, makine, ek, ekip), (adet, saat) in sorted(
            birlesik.items(), key=lambda x: (mdl.malzemeler[x[0][1]].llc, x[0][1], x[0][0])):
        m = mdl.malzemeler[kod]
        kum[kod] += adet
        kk = mdl.kaynaklar.get(kaynak)
        kap = kk.kapasite_sn(gun) / 3600.0 if kk else 0.0
        satirlar.append([gun, kod, m.aciklama[:55], m.grup or "", kaynak,
                         kk.aciklama[:46] if kk else "",
                         mdl.kaynaklar[kaynak].tip if kaynak in mdl.kaynaklar else "",
                         makine or "", ek, m.gerekli_kisi, ekip,
                         round(adet, 3), round(saat, 2), round(kap, 2),
                         round(saat / kap * 100.0, 1) if kap else 0.0, round(kum[kod], 3)])
    tablo(ws, ["Tarih", "Malzeme", "Aciklama", "Kapasite Grubu", "Kaynak",
               "Kaynak / Kisi Adi", "Kaynak Tipi", "Makine", "Kisi Grubu",
               "Gerekli Kisi", "Ekip (diger kisiler)", "Miktar",
               "Yuk (saat)", "Kaynak Kapasitesi (saat)", "Doluluk %", "Kumulatif Uretim"],
          satirlar, [12, 20, 34, 26, 15, 40, 11, 13, 16, 12, 30, 12, 11, 18, 11, 16],
          {"Miktar": "#,##0.###", "Yuk (saat)": "#,##0.00", "Gerekli Kisi": "0",
           "Kaynak Kapasitesi (saat)": "#,##0.00",
           "Doluluk %": "#,##0.0", "Kumulatif Uretim": "#,##0.###"})

    # ---------------- IS EMIRLERI ---------------- #
    # Ayni urun, ayni hafta, ayni kaynak (ve ayni makine) uzerinde devam eden
    # uretim TEK is emrinde toplanir. Kaynak degisirse ya da hafta donerse yeni
    # is emri acilir.
    ws = wb.create_sheet("IsEmirleri")
    emir = OrderedDict()
    for gun, kod, kaynak, makine, adet, saat, talep, isno, ek, ekip in mdl.uretim_detay:
        if adet <= EPS:
            continue
        yil, hft, _ = gun.isocalendar()
        a = (yil, hft, kod, kaynak, makine or "", ek or "", ekip or "")
        if a not in emir:
            emir[a] = {"ilk": gun, "son": gun, "gun": set(), "adet": 0.0,
                       "saat": 0.0, "talep": OrderedDict()}
        e = emir[a]
        e["ilk"] = min(e["ilk"], gun)
        e["son"] = max(e["son"], gun)
        e["gun"].add(gun)
        e["adet"] += adet
        e["saat"] += saat
        e["talep"][talep] = None

    satirlar = []
    for n, ((yil, hft, kod, kaynak, makine, ek, ekip), e) in enumerate(
            sorted(emir.items(), key=lambda x: (x[1]["ilk"], x[0][2], x[0][3])), start=1):
        m_ = mdl.malzemeler[kod]
        kk = mdl.kaynaklar.get(kaynak)
        sip = list(e["talep"])
        satirlar.append([
            "IE-%d-H%02d-%04d" % (yil, hft, n), "%d-H%02d" % (yil, hft),
            kod, m_.aciklama[:48], m_.birim,
            m_.grup or "", kaynak, kk.aciklama[:40] if kk else "", kk.tip if kk else "",
            makine or "", ek or "", m_.gerekli_kisi, ekip or "",
            e["ilk"], e["son"], len(e["gun"]),
            round(e["adet"], 3), round(e["saat"], 2),
            round(e["adet"] / len(e["gun"]), 1),
            ", ".join(sip[:3]) + (" +%d" % (len(sip) - 3) if len(sip) > 3 else "")])
    son = tablo(ws, ["Is Emri", "Hafta", "Urun", "Aciklama", "Birim", "Kapasite Grubu",
                     "Kaynak", "Kaynak / Kisi Adi", "Tip", "Makine", "Kisi Grubu",
                     "Gerekli Kisi", "Ekip (diger kisiler)",
                     "Baslangic", "Bitis", "Gun", "Miktar", "Yuk (saat)",
                     "Gunluk ort. adet", "Tetikleyen siparis"],
                satirlar, [18, 11, 20, 42, 7, 28, 16, 38, 10, 12, 16, 12, 30,
                           12, 12, 7, 14, 12, 15, 34],
                {"Miktar": "#,##0.###", "Yuk (saat)": "#,##0.00", "Gerekli Kisi": "0",
                 "Gunluk ort. adet": "#,##0.#"})
    ws.cell(row=son + 2, column=1,
            value=("%d is emri. Ayni urun + ayni kaynak + ayni makine, ayni hafta icinde "
                   "tek emirde toplandi; kaynak degisirse veya hafta donerse yeni emir acilir. "
                   "'Gerekli Kisi' > 1 ise is, Kaynak + Ekip kolonundaki kisiler tarafindan "
                   "AYNI ANDA yapilir. 'Yuk (saat)' isin SURESIDIR (ekibin her uyesi bu kadar "
                   "mesgul olur); toplam adam-saat = Yuk x Gerekli Kisi."
                   % len(satirlar))).font = Font(italic=True, size=9, color="808080")

    # ---------------- KALIP -> MAKINE ATAMA ---------------- #
    if any(k.makineler for k in mdl.kaynaklar.values()):
        ws = wb.create_sheet("KalipMakineAtama")
        # kalip bazinda makine kullanim donemleri
        donem = OrderedDict()
        for gun, kod, kaynak, makine, adet, saat, talep, isno, ek, ekip in mdl.uretim_detay:
            if not makine:
                continue
            a = (kaynak, makine)
            if a not in donem:
                donem[a] = [gun, gun, 0.0, 0.0, set()]
            d = donem[a]
            d[0] = min(d[0], gun)
            d[1] = max(d[1], gun)
            d[2] += adet
            d[3] += saat
            d[4].add(gun)
        satirlar = []
        for (kalip, makine), (ilk, son, adet, saat, gunler) in sorted(
                donem.items(), key=lambda x: (x[0][0], x[1][0])):
            satirlar.append([kalip, mdl.kaynaklar[kalip].aciklama[:40] if kalip in mdl.kaynaklar
                             else "", makine, ilk, son, len(gunler), round(adet, 3),
                             round(saat, 2)])
        son_satir = tablo(ws, ["Kalip", "Kalip Aciklamasi", "Makine", "Ilk Gun", "Son Gun",
                               "Calisilan Gun", "Uretim (adet)", "Toplam Saat"],
                          satirlar, [20, 38, 16, 13, 13, 15, 16, 14],
                          {"Uretim (adet)": "#,##0.###", "Toplam Saat": "#,##0.00"})

        ws.cell(row=son_satir + 2, column=1, value="KALIP TASIMA GECMISI").font = \
            Font(bold=True, size=11)
        tasima = [[g, kalip, eski or "(ilk baglanti)", yeni_m]
                  for g, kalip, yeni_m, eski in mdl.kalip_gecmisi]
        son_satir = tablo(ws, ["Tarih", "Kalip", "Onceki Makine", "Yeni Makine"],
                          tasima, None, None, baslik_satiri=son_satir + 3)
        sayac = defaultdict(int)
        for g, kalip, yeni_m, eski in mdl.kalip_gecmisi:
            if eski:
                sayac[kalip] += 1
        ws.cell(row=son_satir + 2, column=1,
                value=("Toplam %d kalip tasimasi. Kalip bir makineye baglandiginda orada kalir; "
                       "sadece o makine yer vermediginde tasinir. "
                       "Parametreler > 'Minimum Kalip Kalis (gun)' ile zorlanabilir."
                       % sum(sayac.values()))).font = Font(italic=True, size=9, color="808080")

    # ---------------- GERI PLAN ---------------- #
    ws = wb.create_sheet("GeriPlan")
    satirlar = []
    for kod in sorted(mdl.geri_uretim, key=lambda k: mdl.malzemeler[k].llc):
        m = mdl.malzemeler[kod]
        for gun in sorted(mdl.geri_uretim[kod]):
            q = mdl.geri_uretim[kod][gun]
            satirlar.append([gun, kod, m.kaynak, round(q, 3), round(q * m.sn / 3600.0, 2),
                             "GECMISTE" if gun < mdl.baslangic else ""])
    tablo(ws, ["Tarih", "Malzeme", "Kaynak", "Miktar", "Yuk (saat)", "Uyari"],
          satirlar, [12, 22, 13, 12, 11, 12],
          {"Miktar": "#,##0.###", "Yuk (saat)": "#,##0.00"})

    # ---------------- KAPASITE YUKU ---------------- #
    ws = wb.create_sheet("KapasiteYuku")
    satirlar = []
    ufuk_gun = [g for g in mdl.gun_listesi((mdl.son_gun - mdl.baslangic).days + 1)]
    for kod, kaynak in mdl.kaynaklar.items():
        for gun in ufuk_gun:
            kap = kaynak.kapasite_sn(gun) / 3600.0
            yuk = mdl.yuk[kod].get(gun, 0.0) / 3600.0
            if kap <= 0 and yuk <= 0:
                continue
            satirlar.append([gun, gun.strftime("%a"), kod, kaynak.aciklama[:42],
                             kaynak.tip, round(kap, 2), round(yuk, 2),
                             round(kap - yuk, 2), round(yuk / kap * 100.0, 1) if kap else 0.0])
    son = tablo(ws, ["Tarih", "Gun", "Kaynak", "Kaynak / Kisi Adi", "Tip",
                     "Kapasite (saat)", "Yuk (saat)", "Bos (saat)", "Doluluk %"],
                satirlar, [12, 7, 14, 40, 10, 14, 12, 12, 11],
                {"Kapasite (saat)": "#,##0.00", "Yuk (saat)": "#,##0.00",
                 "Bos (saat)": "#,##0.00", "Doluluk %": "#,##0.0"})
    if son > 1:
        ws.conditional_formatting.add("I2:I%d" % son, ColorScaleRule(
            start_type="num", start_value=0, start_color="FFFFFF",
            mid_type="num", mid_value=80, mid_color="FFEB9C",
            end_type="num", end_value=100, end_color="63BE7B"))

    # ---------------- MRP SONUC (ERP formati) ---------------- #
    ws = wb.create_sheet("MRPSonuc")
    satirlar = []
    for kod in sorted(mdl.malzemeler, key=lambda k: mdl.malzemeler[k].llc):
        m = mdl.malzemeler[kod]
        satirlar.append([kod, "", "Aciklama : " + m.aciklama, "", "", None])
        kum = m.stok
        satirlar.append([kod, m.aciklama[:60], "Baslangic Stogu", round(m.stok, 3),
                         round(kum, 3), mdl.baslangic])
        hrk = sorted(mdl.hareket.get(kod, []), key=lambda x: (x[0], -x[3]))
        for gun, tip, acik, q in hrk:
            kum += q
            satirlar.append([kod, m.aciklama[:60], "%-18s %s" % (tip, acik),
                             round(q, 3), round(kum, 3), gun])
    son = tablo(ws, ["Malzeme", "Malzeme Aciklamasi", "Belge", "Miktar",
                     "Kumulatif Stok", "Tarih"],
                satirlar, [22, 45, 46, 14, 16, 13],
                {"Miktar": "#,##0.###", "Kumulatif Stok": "#,##0.###"})
    ws.conditional_formatting.add("E2:E%d" % son,
                                  CellIsRule(operator="lessThan", formula=["0"], fill=KOTU))

    # ---------------- HAMMADDE PLANI ---------------- #
    ws = wb.create_sheet("HammaddePlani")
    satirlar = []
    for kod, m in mdl.malzemeler.items():
        if m.tip != "SATINALMA":
            continue
        girisler = defaultdict(float)
        for a in mdl.acik:
            if a["malzeme"] == kod:
                girisler[a["tarih"] or mdl.baslangic] += a["miktar"]
        stok = m.stok
        gunler = sorted(set(list(mdl.tuketim[kod]) + list(girisler) + list(mdl.satinalma_ileri[kod])))
        satirlar.append([mdl.baslangic, kod, m.aciklama[:55], m.birim, "Baslangic Stogu",
                         round(m.stok, 3), round(stok, 3), ""])
        for gun in gunler:
            giris = girisler.get(gun, 0.0)
            if giris:
                stok += giris
                satirlar.append([gun, kod, m.aciklama[:55], m.birim, "Tedarik / Giris",
                                 round(giris, 3), round(stok, 3), ""])
            cik = mdl.tuketim[kod].get(gun, 0.0)
            if cik:
                stok -= cik
                acik = mdl.satinalma_ileri[kod].get(gun, 0.0)
                satirlar.append([gun, kod, m.aciklama[:55], m.birim, "Uretim Tuketimi",
                                 round(-cik, 3), round(stok, 3),
                                 "ACIK %.3f - satin alinmali" % acik if acik > EPS else ""])
    son = tablo(ws, ["Tarih", "Malzeme", "Aciklama", "Birim", "Hareket",
                     "Miktar", "Kalan Stok", "Uyari"],
                satirlar, [12, 22, 40, 8, 18, 14, 14, 34],
                {"Miktar": "#,##0.###", "Kalan Stok": "#,##0.###"})
    if son > 1:
        ws.conditional_formatting.add("G2:G%d" % son,
                                      CellIsRule(operator="lessThan", formula=["0"], fill=KOTU))

    # ---------------- SENARYO ANALIZI ---------------- #
    if getattr(mdl, "senaryolar", None):
        ws = wb.create_sheet("SenaryoAnalizi")
        satirlar = [[s["tur"] + 1, s["kurulum"], s["kaynak"], s["paralel"], s["doluluk"],
                     s["gunluk_adet"], s["son_gun"], s["max_gecikme"], s["gec_kalem"], s["durum"]]
                    for s in mdl.senaryolar]
        son = tablo(ws, ["Adim", "Kurulum (kaynak x paralel)", "Darbogaz", "Darbogaz Paralel",
                         "Darbogaz Doluluk %", "Darbogaz Gunluk Kapasite (adet)",
                         "Son Uretim Gunu", "En Buyuk Gecikme (gun)", "Geciken Kalem", "Durum"],
                    satirlar, [6, 46, 14, 16, 18, 26, 17, 20, 14, 24],
                    {"Darbogaz Doluluk %": "#,##0.0",
                     "Darbogaz Gunluk Kapasite (adet)": "#,##0",
                     "En Buyuk Gecikme (gun)": "#,##0"})
        ws.conditional_formatting.add("J2:J%d" % son,
                                      CellIsRule(operator="equal", formula=['"TUM SIPARISLER ZAMANINDA"'],
                                                 fill=IYI))
        ws.cell(row=son + 2, column=1,
                value=("Her adimda o anki darbogaza +1 paralel tezgah/kisi eklenip plan yeniden "
                       "calistirildi. Darbogaz her adimda degisebilir.")).font = \
            Font(italic=True, size=9, color="808080")

    # ---------------- KONTROL ---------------- #
    ws = wb.create_sheet("Kontrol")
    tablo(ws, ["Kontrol", "Sonuc", "Detay"], mdl.kontroller, [46, 12, 70])
    ws.conditional_formatting.add("B2:B%d" % (len(mdl.kontroller) + 1),
                                  CellIsRule(operator="equal", formula=['"GECTI"'], fill=IYI))
    ws.conditional_formatting.add("B2:B%d" % (len(mdl.kontroller) + 1),
                                  CellIsRule(operator="equal", formula=['"HATA"'], fill=KOTU))

    wb.save(yol)


# ====================================================================== #
#  DOGRULAMA
# ====================================================================== #

def kontrol_et(mdl):
    k = []

    # 1) kapasite asimi
    asim = []
    for kod, kaynak in mdl.kaynaklar.items():
        for gun, sn in mdl.yuk[kod].items():
            kap = kaynak.kapasite_sn(gun)
            if sn > kap + 1e-3:
                asim.append("%s %s: %.2f > %.2f saat" % (kod, gun, sn / 3600, kap / 3600))
    k.append(["Hicbir gunde kapasite asilmadi", "GECTI" if not asim else "HATA",
              "; ".join(asim[:3]) if asim else "%d kaynak kontrol edildi" % len(mdl.kaynaklar)])

    # 2) calisilmayan gunde uretim
    hatali = []
    for kod, kaynak in mdl.kaynaklar.items():
        for gun, sn in mdl.yuk[kod].items():
            if sn > EPS and not kaynak.calisir(gun):
                hatali.append("%s %s" % (kod, gun))
    k.append(["Tatil / calisilmayan gunde uretim yok", "GECTI" if not hatali else "HATA",
              "; ".join(hatali[:3]) if hatali else "Hafta sonu ve tatiller bos"])

    # 3) stok eksiye dusmedi (uretim malzemeleri)
    eksi = []
    for kod, m in mdl.malzemeler.items():
        if m.tip == "SATINALMA" and not mdl.satinalma_kisit:
            continue
        kum = m.stok
        for gun, tip, acik, q in sorted(mdl.hareket.get(kod, []), key=lambda x: (x[0], -x[3])):
            kum += q
            if kum < -1e-3:
                eksi.append("%s %s -> %.2f" % (kod, gun, kum))
                break
    k.append(["Uretim malzemelerinde negatif stok yok", "GECTI" if not eksi else "HATA",
              "; ".join(eksi[:3]) if eksi else "Tum seviyeler pozitif"])

    # 4) malzeme dengesi: uretim = is miktari
    fark = []
    for kod, isler in mdl.isler.items():
        planlanan = sum(i.miktar for i in isler)
        uretilen = sum(mdl.uretim[kod].values())
        if abs(planlanan - uretilen) > 1e-3:
            fark.append("%s: plan %.1f / uretim %.1f" % (kod, planlanan, uretilen))
    k.append(["Planli emirlerin tamami uretildi", "GECTI" if not fark else "UYARI",
              "; ".join(fark[:3]) if fark else "Plan = uretim"])

    # 5) talep dengesi
    tk = sum(t["miktar"] for t in mdl.talepler)
    ts = sum(q for _, _, q in mdl.teslim)
    k.append(["Sevkiyat = talep", "GECTI" if abs(tk - ts) < 1e-3 else "UYARI",
              "talep %.1f / sevk %.1f" % (tk, ts)])

    # 6) paralellik: seviyeler ust uste biniyor mu?
    araliklar = {}
    for kod in mdl.uretim:
        g = sorted(mdl.uretim[kod])
        if g:
            araliklar[kod] = (g[0], g[-1])
    ortusme = []
    for ust, m in mdl.malzemeler.items():
        for alt, _ in m.cocuklar:
            if ust in araliklar and alt in araliklar:
                a1, b1 = araliklar[ust]
                a2, b2 = araliklar[alt]
                gun = (min(b1, b2) - max(a1, a2)).days + 1
                if gun > 0:
                    ortusme.append("%s/%s %d gun" % (ust, alt, gun))
    k.append(["Seviyeler paralel calisiyor (ust uste biniyor)",
              "GECTI" if ortusme else "UYARI",
              "; ".join(ortusme) if ortusme else "Ortusme bulunamadi"])

    # 7) aktarma kurali: bir gun icinde uretileni ayni gun tuketmek yok
    ihlal = []
    gecikme = mdl.aktarma_gecikmesi()
    ayni_gun = 0
    for alt, m in mdl.malzemeler.items():
        if m.tip != "URETIM" or not m.ebeveynler:
            continue
        kum_tuketim = 0.0
        gunler = sorted(set(list(mdl.giris[alt]) + list(mdl.tuketim[alt])))
        for gun in gunler:
            kum_tuketim += mdl.tuketim[alt].get(gun, 0.0)
            hazir = m.stok + sum(q for g2, q in mdl.giris[alt].items() if g2 <= gun)
            if kum_tuketim > hazir + 1e-3:
                ihlal.append("%s %s (tuketim %.1f > hazir %.1f)" % (alt, gun, kum_tuketim, hazir))
                break
            if gecikme and mdl.uretim[alt].get(gun, 0.0) > EPS and \
                    mdl.tuketim[alt].get(gun, 0.0) > EPS:
                # ayni gun hem uretim hem tuketim olabilir; onemli olan tuketilenin
                # o gunden ONCE hazir olmasi - yukaridaki kosul bunu zaten garanti eder
                ayni_gun += 1
    k.append(["Aktarma kurali korundu (uretilen ertesi gun kullanilabilir)",
              "GECTI" if not ihlal else "HATA",
              "; ".join(ihlal[:3]) if ihlal
              else "Kumulatif tuketim her gun <= o gune kadar hazir olan miktar"])

    # 8) alternatif kaynak onceligi: dusuk oncelikli kaynak ancak yuksek
    #    oncelikli DOLDUGUNDA devreye girmis olmali
    ihlal8 = []
    kullanim = defaultdict(set)          # (gun, malzeme) -> kullanilan kaynaklar
    for gun, kod, kaynak, makine, adet, saat, talep, isno, ek, ekip in mdl.uretim_detay:
        if adet > EPS:
            kullanim[(gun, kod)].add(kaynak)
    for (gun, kod), kaynaklar in kullanim.items():
        alts = mdl.malzemeler[kod].alternatifler
        sira = {a.kaynak: n for n, a in enumerate(alts)}
        if not kaynaklar or max(sira.get(x, 0) for x in kaynaklar) == 0:
            continue
        en_dusuk = max(sira.get(x, 0) for x in kaynaklar)
        kalan = mdl.kalan_kapasite.get(gun, {})
        for a in alts[:en_dusuk]:
            # Oncelikli kaynak o gun bu urun icin zaten acilmissa oncelige
            # uyulmustur; gun sonunda bos kalmasi malzeme/bilesen beklemesinden
            # kaynaklanir (dagitim gun icinde parca parca yapiliyor).
            if a.kaynak in kaynaklar:
                continue
            # daha oncelikli kaynak gun sonunda hala bos mu?  (bos ise ihlal)
            if kalan.get(a.kaynak, 0.0) > a.sn + 1e-3:
                # Gun ici is degisimi siniri o kaynagi bilerek kapatmis olabilir.
                # Iki durumda ihlal sayilmaz:
                #  a) kaynak kotasini doldurmus ve bu urune hic gecmemis,
                #  b) kotasini ASMIS - yani gun icinde bir noktada bloke olup
                #     ancak gecikmis isler icin acilan serbest fazda calismis.
                if mdl.gunluk_is_cesidi > 0:
                    yap = mdl.gun_is_kaydi.get(gun, {}).get(a.kaynak)
                    if yap and ((kod not in yap and len(yap) >= mdl.gunluk_is_cesidi)
                                or len(yap) > mdl.gunluk_is_cesidi):
                        break
                # Is basina kaynak siniri o hafta icin dolmus olabilir: daha
                # oncelikli kaynak bos dursa da bilerek acilmamistir.
                # Is basina kaynak siniri aktifken is, hafta basinda secilen
                # kaynak kumesinde kalir; kumenin disindaki bir kaynak bos dursa
                # da bilerek acilmaz (is kalabaliga bolunmesin diye). Oncelik
                # ihlali ancak KUME ICINDEKI kaynaklar arasinda anlamlidir.
                if mdl.is_basina_kaynak > 0:
                    hk = mdl.hafta_kaynak.get((gun.isocalendar()[:2], kod))
                    if hk and a.kaynak not in hk:
                        break
                bos_makine = True
                kk = mdl.kaynaklar.get(a.kaynak)
                if kk and kk.makineler:
                    # kalibin butun makineleri dolmus olabilir - o zaman ihlal degil
                    bos_makine = any(kalan.get(mk, 0.0) > a.sn * c + 1e-3
                                     for mk, c in kk.makineler)
                if bos_makine:
                    ihlal8.append("%s %s: %s bos dururken dusuk oncelikli kaynak kullanildi"
                                  % (gun, kod, a.kaynak))
                break
    k.append(["Alternatif kaynak onceligine uyuldu (gun ici is ve kaynak sinirlari dikkate alinarak)",
              "GECTI" if not ihlal8 else "UYARI",
              "; ".join(ihlal8[:3]) if ihlal8
              else "Dusuk oncelikli kaynak sadece ustu doldugunda devreye girdi"])

    # 9) kalip yalnizca tanimli makinelerinde calisti mi
    ihlal9 = []
    for gun, kod, kaynak, makine, adet, saat, talep, isno, ek, ekip in mdl.uretim_detay:
        kk = mdl.kaynaklar.get(kaynak)
        if makine and kk and makine not in [x for x, _ in kk.makineler]:
            ihlal9.append("%s %s -> %s (tanimli degil)" % (gun, kaynak, makine))
        if kk and kk.makineler and not makine:
            ihlal9.append("%s %s: makineye baglanmadan calisti" % (gun, kaynak))
    kalipli = [k2 for k2, v in mdl.kaynaklar.items() if v.makineler]
    k.append(["Kaliplar yalnizca tanimli makinelerinde calisti",
              "GECTI" if not ihlal9 else "HATA",
              "; ".join(ihlal9[:3]) if ihlal9
              else ("%d kalip kontrol edildi, %d tasima"
                    % (len(kalipli), len([x for x in mdl.kalip_gecmisi if x[3]])))])

    mdl.kontroller = k
    return k


# ====================================================================== #
#  SENARYO: darbogaz kaynakta paralel tezgah sayisi artirilirsa?
# ====================================================================== #

def _doluluk(m2):
    """Kaynak -> doluluk orani (ileri plan calistirildiktan sonra)."""
    out = {}
    gunler = m2.gun_listesi((m2.son_gun - m2.baslangic).days + 1)
    for kod, kaynak in m2.kaynaklar.items():
        mevcut = sum(kaynak.kapasite_sn(g) for g in gunler)
        out[kod] = (sum(m2.yuk[kod].values()) / mevcut) if mevcut else 0.0
    return out


def senaryo_calistir(girdi, adim=12):
    """
    Kademeli kapasite artirimi: her turda o anki darbogaza +1 paralel tezgah/kisi
    eklenir ve sonuc olculur. Boylece 'nereye ne kadar yatirim' sorusu cevaplanir.
    """
    sonuclar = []
    ilave = defaultdict(int)
    for tur in range(adim):
        m2 = Model(girdi)
        for k, n in ilave.items():
            m2.kaynaklar[k].paralel += n
        m2.patlat()
        m2.ileri_planla()
        gec = [t for t in m2.talep_sonuc if t["karsilanan_tarih"] and t["karsilanan_tarih"] > t["tarih"]]
        yok = [t for t in m2.talep_sonuc if not t["karsilanan_tarih"]]
        maks = max([(t["karsilanan_tarih"] - t["tarih"]).days for t in gec] or [0])
        dol = _doluluk(m2)
        darbogaz = max(dol, key=lambda k: dol[k])

        adet_gun = 0.0
        for kod, mm in m2.malzemeler.items():
            if mm.kaynak == darbogaz and mm.sn > 0:
                adet_gun = m2.kaynaklar[darbogaz].saat_gun * 3600.0 * \
                    m2.kaynaklar[darbogaz].paralel / mm.sn
                break
        sonuclar.append({
            "tur": tur,
            "kurulum": ", ".join("%s x%d" % (k, m2.kaynaklar[k].paralel)
                                 for k in m2.kaynaklar),
            "kaynak": darbogaz,
            "paralel": m2.kaynaklar[darbogaz].paralel,
            "doluluk": round(dol[darbogaz] * 100, 1),
            "gunluk_adet": round(adet_gun),
            "son_gun": m2.son_gun, "max_gecikme": maks,
            "gec_kalem": len(gec) + len(yok),
            "durum": "TUM SIPARISLER ZAMANINDA" if not gec and not yok
                     else ("%d kalem gecikiyor" % (len(gec) + len(yok))),
        })
        if not gec and not yok:
            break
        if tur and sonuclar[-1]["son_gun"] == sonuclar[-2]["son_gun"] and \
                sonuclar[-1]["kaynak"] == sonuclar[-2]["kaynak"]:
            break                                  # ilerleme yok, dur
        ilave[darbogaz] += 1
    return sonuclar


# ====================================================================== #
def main():
    girdi = sys.argv[1] if len(sys.argv) > 1 else "MRP_Girdi.xlsx"
    cikti = sys.argv[2] if len(sys.argv) > 2 else "MRP_Sonuc.xlsx"

    print("Girdi okunuyor  :", girdi)
    mdl = Model(girdi)
    print("  malzeme %d | kaynak %d | talep %d" %
          (len(mdl.malzemeler), len(mdl.kaynaklar), len(mdl.talepler)))

    print("MRP patlatiliyor...")
    mdl.patlat()
    for kod in sorted(mdl.isler, key=lambda k: mdl.malzemeler[k].llc):
        print("  %-22s %2d emir / %12.1f %s" % (
            kod, len(mdl.isler[kod]), sum(i.miktar for i in mdl.isler[kod]),
            mdl.malzemeler[kod].birim))

    print("Geri planlama (teslim tarihinden geriye)...")
    mdl.geri_planla()

    print("Ileri planlama (bugunden ileri, paralel akis)...")
    mdl.ileri_planla()

    print("Kontroller...")
    for ad, sonuc, detay in kontrol_et(mdl):
        print("  [%-6s] %s" % (sonuc, ad))

    mdl.senaryolar = []
    if (_s(mdl.p.get("senaryo analizi")) or "EVET").upper() == "EVET" and mdl.kaynaklar:
        print("Senaryo analizi (kademeli kapasite artirimi)...")
        mdl.senaryolar = senaryo_calistir(girdi)
        for s in mdl.senaryolar:
            print("  %-46s darbogaz %-11s %%%5.1f -> son gun %s (%s)"
                  % (s["kurulum"], s["kaynak"], s["doluluk"], s["son_gun"], s["durum"]))

    rapor_yaz(mdl, cikti)
    print("Cikti yazildi   :", cikti)

    gec = [t for t in mdl.talep_sonuc if t["karsilanan_tarih"] and t["karsilanan_tarih"] > t["tarih"]]
    yok = [t for t in mdl.talep_sonuc if not t["karsilanan_tarih"]]
    print("\nSONUC: %d kalemden %d'i zamaninda, %d'i gecikmeli, %d'i ufuk icinde karsilanamadi."
          % (len(mdl.talep_sonuc), len(mdl.talep_sonuc) - len(gec) - len(yok), len(gec), len(yok)))


if __name__ == "__main__":
    main()

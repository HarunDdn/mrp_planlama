# -*- coding: utf-8 -*-
"""
HAM VERI -> PLAN GIRDISI DONUSTURUCUSU
======================================

MRP_Girdi.xlsx icindeki ERP ham sayfalarini okur ve mrp_planlama.py'nin
bekledigi sablon sayfalarini iceren yeni bir dosya uretir.

    python veri_aktar.py MRP_Girdi.xlsx MRP_Plan_Girdi.xlsx

Okudugu ham sayfalar
--------------------
  URUN AGACI SEVIYELI : Seviye | Malzeme | Aciklama | Miktar | Birim |
                        IS MERKEZI / KAPASITE GRUBU | CALISAN PERSONEL SAYISI |
                        ROTA SURESI (sn) | FIRINALAMA
  PERSONELLISTESI     : Gorunen Ad | Sicil No | Gorev | KAPASITE GRUPLARI
  PAKETLEME           : KAPASITE GRUBU | IS MERKEZI (virgullu paketci kodlari)
  STOK                : SAP stok dokumu  -> Kullanilabilir Stok. Plana girmeyecek
                        depolar ayarlar.json > plan.haric_depolar listesindedir.
  SIPARIS             : satis siparisleri -> Acik Miktar

Is merkezi alaninin cozumlenmesi
--------------------------------
  "PKT-GRUP-07"              -> PAKETLEME'deki paketciler, listedeki sira = oncelik
  "MONTAJ" / "BILEZIK" / ... -> PERSONELLISTESI'nde o gruptaki kisiler, sira = oncelik
  "FIRIN1;FIRIN2;..."        -> firinlar, sira = oncelik
  "KM25 (9), KM28 (5)"       -> malzeme bir KALIP'tir; parantezli liste onayli
                                makineler, soldan saga oncelik (parantez = gecmis
                                kullanim adedi)
  bos                        -> kapasitesi olmayan ara adim (KAYNAK-TANIMSIZ)

CALISAN PERSONEL SAYISI = o is icin AYNI ANDA calismasi gereken kisi sayisi.
Rota suresi 1 adedin ekip tarafindan kac saniyede bitirildigini soyler; sure
carpilmaz, bunun yerine plan motoru ayni gun ayni isi yapmak uzere kapasite
grubundan o kadar kisiyi birlikte rezerve eder (Rotalar > "Gerekli Kisi").
Grupta yeterli kisi yoksa motor ekibi kuculttugu kadar sureyi carpar, boylece
toplam adam-saat her durumda korunur.
"""

import datetime as dt
import json
import os
import re
import sys
import unicodedata
from collections import OrderedDict, defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------- #
#  AYARLAR  (Kaynaklar sayfasinda sonradan elle degistirilebilir)
# --------------------------------------------------------------------- #
KISI_SAAT, KISI_GUN = 9.0, [1, 1, 1, 1, 1, 0, 0]       # 9 sa/gun, Pzt-Cum
MAKINE_SAAT, MAKINE_GUN = 24.0, [1, 1, 1, 1, 1, 1, 0]  # 3x8 vardiya donuyor, Pzt-Cmt
KALIP_SAAT, KALIP_GUN = 24.0, [1, 1, 1, 1, 1, 1, 0]    # kalip makineyle ayni takvim
FIRIN_SAAT, FIRIN_GUN = 24.0, [1, 1, 1, 1, 1, 1, 0]    # firinlar da 24 sa, operator gerekmiyor
# Enjeksiyon operatorleri: 8 sa/gun, cumartesi dahil (vardiya donerek makineyi 24 sa calistirirlar)
OPERATOR_GRUP = "ENJEKSIYON"
OPERATOR_SAAT, OPERATOR_GUN = 8.0, [1, 1, 1, 1, 1, 1, 0]
OPERATOR_HAVUZ = "ENJ-OPERATOR"        # enjeksiyon makinelerinin cektigi operator havuzu
OPERATOR_MAKINE = 1.0                  # 1 operator = 1 makine (1 sa makine -> 1 sa operator)
TANIMSIZ = "KAYNAK-TANIMSIZ"       # is merkezi girilmemis adimlar icin
TANIMSIZ_SN = 0.001                # kapasite tuketmesin

AGAC_SAYFALARI = ["URUN AGACI REVIZYONLU", "URUN AGACI SEVIYELI"]   # oncelik sirasiyla
AYAR_DOSYA = "ayarlar.json"          # HTML cizelgesinden kaydedilen ayarlar
KAPANIS_DOSYA = "kapanislar.json"    # kaynak kapatmalari (izin / ariza / bakim)
HARIC_DEPOLAR = ["400", "500"]       # stoklari plana ALINMAYAN depolar

VARSAYILAN = {
    "kisi":      {"saat": KISI_SAAT,     "gunler": KISI_GUN},
    "operator":  {"saat": OPERATOR_SAAT, "gunler": OPERATOR_GUN},
    "makine":    {"saat": MAKINE_SAAT,   "gunler": MAKINE_GUN},
    "kalip":     {"saat": KALIP_SAAT,    "gunler": KALIP_GUN},
    "firin":     {"saat": FIRIN_SAAT,    "gunler": FIRIN_GUN},
}


def _bicim(x):
    """1234567.25 -> '1.234.567,3' (binlik nokta, ondalik virgul)."""
    tam = "{:,.0f}".format(abs(x)).replace(",", ".")
    ond = abs(x) - int(abs(x))
    return ("-" if x < 0 else "") + tam + (",%d" % round(ond * 10) if ond >= 0.05 else "")


def kapanislari_oku(klasor):
    """
    kapanislar.json -> [{kaynak, bas, bit, tur, oran, not}]
    Cizelge HTML'indeki "Kaynak kapatma" ekrani bu dosyayi yazar.
    """
    yol = os.path.join(klasor, KAPANIS_DOSYA)
    if not os.path.exists(yol):
        return []
    try:
        with open(yol, encoding="utf-8") as f:
            d = json.load(f) or {}
    except Exception as e:
        print("UYARI: %s okunamadi (%s) - kapatmalar yok sayildi" % (KAPANIS_DOSYA, e))
        return []
    lst = d.get("kapanislar") if isinstance(d, dict) else d
    return [x for x in (lst or []) if isinstance(x, dict) and x.get("kaynak")]


def ayarlari_oku(klasor="."):
    """ayarlar.json varsa oku; yoksa bos don. HTML cizelgesi bu dosyayi yazar."""
    yol = os.path.join(klasor, AYAR_DOSYA)
    if not os.path.exists(yol):
        return {}
    try:
        with open(yol, encoding="utf-8") as f:
            return json.load(f) or {}
    except (OSError, ValueError) as e:
        print("UYARI: %s okunamadi (%s) - varsayilanlar kullanildi" % (AYAR_DOSYA, e))
        return {}


MAKINE_RE = re.compile(r'([A-Za-z][A-Za-z0-9_\-]*)\s*\((\d+)\)')


# --------------------------------------------------------------------- #
def _s(v):
    return "" if v is None else str(v).strip()


PARTI_RE = re.compile(r"([\d.,]+)\s*(?:AD|ADET|PCS)?\s*[/;|-]?\s*([\d.,]+)\s*(?:H|SA|SAAT|S)\b",
                      re.IGNORECASE)


def parti_coz(v):
    """'20000 AD/ 20 H' -> (20000.0, 20.0)   (parti adedi, parti suresi saat)."""
    m = PARTI_RE.search(_s(v))
    if not m:
        return None
    qs = m.group(1)
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", qs):          # 3.440 -> 3440 (binlik nokta)
        qs = qs.replace(".", "")
    q = float(qs.replace(",", "."))
    t = float(m.group(2).replace(",", "."))
    return (q, t)


def _f(v, vars=0.0):
    if isinstance(v, (int, float)):
        return float(v)
    if v is None or not str(v).strip():
        return vars
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return vars


def _d(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = _s(v)
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def slug(s, n=18):
    """Kaynak kodu icin sadelestirilmis ad."""
    s = unicodedata.normalize("NFKD", _s(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^A-Za-z0-9]+", "-", s).strip("-").upper()
    return s[:n] or "X"


def basliklar(ws):
    for r in range(1, min(ws.max_row, 8) + 1):
        vals = [_s(ws.cell(row=r, column=c).value) for c in range(1, ws.max_column + 1)]
        if len([v for v in vals if v]) >= 2:
            return r, vals
    return 1, []


def oku(wb, ad):
    """Sayfayi sozluk listesi olarak dondur."""
    if ad not in wb.sheetnames:
        return []
    ws = wb[ad]
    bs, bas = basliklar(ws)
    out = []
    for r in range(bs + 1, ws.max_row + 1):
        d = {}
        dolu = False
        for c, b in enumerate(bas, start=1):
            if not b:
                continue
            v = ws.cell(row=r, column=c).value
            d[b] = v
            if v is not None and _s(v) != "":
                dolu = True
        if dolu:
            out.append(d)
    return out


def kolon(d, *adlar):
    def n(x):
        x = _s(x).lower()
        for a, b in (("ı", "i"), ("İ", "i"), ("ş", "s"), ("ğ", "g"), ("ü", "u"),
                     ("ö", "o"), ("ç", "c"), (" ", ""), ("_", ""), ("/", ""),
                     ("(", ""), (")", ""), (".", "")):
            x = x.replace(a, b)
        return x
    hedef = {n(a) for a in adlar}
    for k, v in d.items():
        if n(k) in hedef:
            return v
    return None


# ===================================================================== #
class Donustur:
    def __init__(self, girdi, ayar=None):
        self.wb = load_workbook(girdi, data_only=True)
        self.ayar = ayar or {}
        v = dict(VARSAYILAN)
        for k, d in (self.ayar.get("varsayilan") or {}).items():
            if k in v and isinstance(d, dict):
                v[k] = {"saat": float(d.get("saat", v[k]["saat"])),
                        "gunler": list(d.get("gunler", v[k]["gunler"]))}
        self.v = v
        self.op_katsayi = float((self.ayar.get("operator") or {}).get("katsayi", OPERATOR_MAKINE))
        self.op_kisit = bool((self.ayar.get("operator") or {}).get("kisit", True))
        self.kaynak_ayar = self.ayar.get("kaynaklar") or {}
        # Plana alinmayacak depolar (orn. mamul sevk / karantina depolari).
        hd = (self.ayar.get("plan") or {}).get("haric_depolar", HARIC_DEPOLAR)
        if isinstance(hd, str):
            hd = [x.strip() for x in re.split(r"[,;\s]+", hd) if x.strip()]
        self.haric_depolar = {str(x).strip() for x in (hd or [])}
        self.uyari = []
        self.malzeme = OrderedDict()     # kod -> dict
        self.kaynak = OrderedDict()      # kod -> dict
        self.rota = []                   # (malzeme, oncelik, kaynak, sn, gerekli_kisi)
        self.kalipmakine = []            # (kalip, oncelik, makine, carpan)
        self.kalip_parca = {}            # kalip kodu -> ait oldugu ENJ parcasi
        self.agac = []                   # (seviye, malzeme, miktar, fire)
        self.talep = []
        self._personel()
        self._urun_agaci()
        self._kaynak_ayarlarini_uygula()
        self._stok()
        self._siparis()

    # ----------------------------------------------------------------- #
    def ekle_kaynak(self, kod, aciklama, tip, saat, gunler, paralel=1,
                    ek="", ek_kats=""):
        if kod not in self.kaynak:
            self.kaynak[kod] = dict(kod=kod, aciklama=aciklama, tip=tip,
                                    saat=saat, paralel=paralel, gunler=gunler,
                                    ek=ek, ek_kats=ek_kats)
        return kod

    # ----------------------------------------------------------------- #
    def _personel(self):
        """PERSONELLISTESI + PAKETLEME -> grup/kod -> kaynak kodu listesi."""
        self.grup = defaultdict(list)        # kapasite grubu -> [kaynak kodu]
        self.kisi_kodu = {}                  # PAKETLEME kodu -> kaynak kodu
        for r in oku(self.wb, "PERSONELLISTESI"):
            ad = _s(kolon(r, "Görünen Ad", "Gorunen Ad", "Ad"))
            if not ad:
                continue
            sicil = _s(kolon(r, "Sicil No"))
            gorev = _s(kolon(r, "Görev", "Gorev"))
            # kaynak kodu = personel sicil numarasi (onek yok)
            kod = sicil or slug(ad, 10)
            # Bir kisi birden fazla kapasite grubunda olabilir: KAPASITE GRUPLARI,
            # KAPASITE GRUPLARI2, 3, 4 ... kolonlarinin hepsi okunur, ayrica her
            # hucre icinde virgul/noktali virgul ile birden fazla grup yazilabilir.
            gruplar = []
            for anahtar, deger in r.items():
                if "KAPASITE" not in _s(anahtar).upper():
                    continue
                for g in re.split(r"[,;\n]", _s(deger)):
                    g = g.strip()
                    if g and g not in gruplar:
                        gruplar.append(g)
            if OPERATOR_GRUP in gruplar:
                self.ekle_kaynak(kod, "%s - %s" % (ad, gorev), "KISI",
                                 self.v["operator"]["saat"], self.v["operator"]["gunler"])
            else:
                self.ekle_kaynak(kod, "%s - %s" % (ad, gorev), "KISI", self.v["kisi"]["saat"], self.v["kisi"]["gunler"])
            for g in gruplar:
                self.grup[g].append(kod)
                if len(g) <= 12 and g.upper() == g.replace(" ", "").upper() and " " not in g:
                    self.kisi_kodu[g] = kod       # paketci kisa kodu (ACAKAL vb.)

        # enjeksiyon operator havuzu: 1 operator = 1 makine
        # Bir enjeksiyon operatoru baska bir kapasite grubunda da yer aliyorsa
        # orada BIREYSEL kaynak olarak planlaniyor demektir; kapasitesi iki kez
        # sayilmasin diye operator havuzundan cikarilir.
        self.havuz_disi = sorted({k for k in self.grup.get(OPERATOR_GRUP, [])
                                  if any(g != OPERATOR_GRUP and k in v
                                         for g, v in self.grup.items())})
        havuz_uye = [k for k in self.grup.get(OPERATOR_GRUP, []) if k not in self.havuz_disi]
        opsay = len(havuz_uye)
        if opsay:
            self.ekle_kaynak(OPERATOR_HAVUZ,
                             "Enjeksiyon operator havuzu - %d kisi x %g sa"
                             % (opsay, self.v["operator"]["saat"]),
                             "KISI", OPERATOR_SAAT, OPERATOR_GUN, paralel=opsay)
            if self.havuz_disi:
                self.uyari.append(
                    "%d enjeksiyon operatoru baska kapasite grubunda da gorevli; "
                    "cift sayilmamasi icin operator havuzundan cikarildi (%s)"
                    % (len(self.havuz_disi), ", ".join(self.havuz_disi)))
        self.operator_sayisi = opsay
        self.havuz_uye = havuz_uye

        self.pkt = {}
        for r in oku(self.wb, "PAKETLEME"):
            g = _s(kolon(r, "KAPASITE GRUBU"))
            if not g:
                continue
            self.pkt[g] = [x.strip() for x in _s(kolon(r, "IS MERKEZI", "İŞ MERKEZİ")).split(",") if x.strip()]

    # ----------------------------------------------------------------- #
    def _grup_adi(self, ham):
        """Is merkezi alanindan kapasite grubu adini cikar (makine listesi ise ENJEKSIYON)."""
        ham = _s(ham)
        if ham in ("0", "0.0", "-", ""):
            return ""
        if MAKINE_RE.findall(ham):
            return OPERATOR_GRUP
        if ";" in ham:
            return "FIRIN"
        return ham

    # ----------------------------------------------------------------- #
    def _is_merkezi(self, kod, ham, sure, personel):
        """
        Bir uretim adiminin kaynak alternatiflerini uret.
        Dondurur: [(oncelik, kaynak_kodu, saniye), ...]
        """
        ham = _s(ham)
        if ham in ("0", "0.0", "-"):          # ERP'de bos yerine 0 yazilmis
            ham = ""
        # Sure kisi sayisiyla CARPILMAZ: personel sayisi ekip buyuklugudur ve
        # Rotalar > "Gerekli Kisi" kolonuyla motora ayrica bildirilir.
        sn = sure
        alt = []

        # 1) enjeksiyon: "KM25 (9), KM28 (5)"  -> malzeme bir KALIP
        ciftler = MAKINE_RE.findall(ham)
        if ciftler:
            kalip = "KLP-" + kod.replace("-R00", "").replace("ENJ", "E")
            self.ekle_kaynak(kalip, "Kalip: %s" % kod, "KALIP", self.v["kalip"]["saat"], self.v["kalip"]["gunler"])
            self.kalip_parca[kalip] = kod
            for i, (mak, _say) in enumerate(ciftler, start=1):
                mk = mak.upper()
                self.ekle_kaynak(mk, "Enjeksiyon makinesi", "MAKINE", self.v["makine"]["saat"], self.v["makine"]["gunler"],
                                 ek=OPERATOR_HAVUZ if (self.operator_sayisi and self.op_kisit) else "",
                                 ek_kats=self.op_katsayi if (self.operator_sayisi and self.op_kisit) else "")
                self.kalipmakine.append((kalip, i, mk, 1.0))
            return [(1, kalip, sn)]

        # 2) firin listesi: "FIRIN1;FIRIN2;..."
        if ";" in ham:
            for i, f in enumerate([x.strip() for x in ham.split(";") if x.strip()], start=1):
                self.ekle_kaynak(f.upper(), "Firin", "MAKINE", self.v["firin"]["saat"], self.v["firin"]["gunler"])
                alt.append((i, f.upper(), sn))
            return alt

        # 3) paketleme grubu
        if ham.upper().startswith("PKT-GRUP"):
            kisiler = self.pkt.get(ham, [])
            if not kisiler:
                self.uyari.append("PAKETLEME sayfasinda '%s' grubu yok - tek kisilik "
                                  "kaynak olusturuldu (%s)" % (ham, kod))
                self.ekle_kaynak(slug(ham), ham, "KISI", self.v["kisi"]["saat"], self.v["kisi"]["gunler"])
                return [(1, slug(ham), sn)]
            for i, kk in enumerate(kisiler, start=1):
                rk = self.kisi_kodu.get(kk)
                if rk is None:
                    rk = self.ekle_kaynak(slug(kk), "Paketci %s" % kk, "KISI", self.v["kisi"]["saat"], self.v["kisi"]["gunler"])
                    self.uyari.append("Paketci kodu '%s' PERSONELLISTESI'nde yok (%s)" % (kk, ham))
                alt.append((i, rk, sn))
            return alt

        # 4) personel havuzu (MONTAJ, BILEZIK, CAPAK ALMA / ZIMPARA ...)
        if ham in self.grup and self.grup[ham]:
            for i, rk in enumerate(self.grup[ham], start=1):
                alt.append((i, rk, sn))
            return alt

        # 5) tanimsiz is merkezi -> tek kisilik kaynak, elle duzeltilsin
        if ham:
            self.uyari.append("'%s' is merkezi PERSONELLISTESI/PAKETLEME'de yok - "
                              "tek kapasiteli kaynak varsayildi (%s)" % (ham, kod))
            self.ekle_kaynak(slug(ham), ham, "KISI", self.v["kisi"]["saat"], self.v["kisi"]["gunler"])
            return [(1, slug(ham), sn)]

        # 6) is merkezi hic yok -> kapasitesiz ara adim
        self.ekle_kaynak(TANIMSIZ, "Is merkezi girilmemis adimlar - kapasite kisiti yok",
                         "DIGER", 24.0, [1] * 7, paralel=999)
        return [(1, TANIMSIZ, TANIMSIZ_SN)]

    # ----------------------------------------------------------------- #
    def _urun_agaci(self):
        # Revizyonlu sayfa (acik "Ust Malzeme" kolonlu) varsa o kullanilir.
        self.agac_sayfa = None
        for ad in AGAC_SAYFALARI:
            satirlar = oku(self.wb, ad)
            if satirlar:
                self.agac_sayfa = ad
                break
        if not self.agac_sayfa:
            raise SystemExit("Urun agaci sayfasi bulunamadi veya bos (%s)."
                             % " / ".join(AGAC_SAYFALARI))
        ham_satir = []
        for r in satirlar:
            kod = _s(kolon(r, "Malzeme Kodu", "Malzeme"))
            if not kod:
                continue
            sev = kolon(r, "Ürün Ağacı Seviyesi", "Urun Agaci Seviyesi", "Seviye")
            if sev is None or _s(sev) == "":
                self.uyari.append("Seviyesi bos satir atlandi: %s" % kod)
                continue
            ham_satir.append((int(_f(sev)), kod, r, _s(kolon(r, "Üst Malzeme", "Ust Malzeme"))))

        # ---- her satirin EBEVEYNI ----
        # 1) "Ust Malzeme" kolonu doluysa ve o kod sayfada gercekten bir satirsa o
        #    kullanilir (satir sirasi karismis olsa bile dogru baglanir);
        # 2) degilse girintiden (ustteki en yakin seviye-1 satiri) bulunur;
        # 3) seviye atlamis ve Ust Malzeme de gecersizse satir atlanir.
        bilinen = {k for _, k, _, _ in ham_satir}
        yigin, onceki_sev = {}, None
        bag = OrderedDict()                  # (ust, alt) -> miktar
        cocuklar = defaultdict(list)         # ust -> [alt] (sayfa sirasiyla)
        kokler = []
        self.agac_duzeltme = []              # (satir, kod, ust_kolon, girinti, karar)
        for n, (sev, kod, r, ust_k) in enumerate(ham_satir, start=1):
            if sev == 0:
                if kod not in kokler:
                    kokler.append(kod)
                yigin = {0: kod}
                onceki_sev = 0
                continue
            girinti = yigin.get(sev - 1)
            atlama = onceki_sev is not None and sev > onceki_sev + 1
            if ust_k and ust_k in bilinen and ust_k != kod:
                ust = ust_k
                if girinti != ust_k:
                    self.agac_duzeltme.append((kod, ust_k, girinti, "Ust Malzeme kolonu kullanildi"))
            elif girinti and not atlama:
                ust = girinti
                if ust_k and ust_k != girinti:
                    self.agac_duzeltme.append((kod, ust_k, girinti,
                                               "Ust Malzeme sayfada yok, girinti kullanildi"))
            else:
                ust = None
                self.agac_duzeltme.append((kod, ust_k, girinti,
                                           "seviye atladi + Ust Malzeme gecersiz: satir atlandi"))
            yigin[sev] = kod
            for d in [x for x in yigin if x > sev]:
                del yigin[d]
            onceki_sev = sev
            if ust is None:
                continue
            miktar = _f(kolon(r, "Miktar"), 1.0) or 1.0
            if (ust, kod) in bag:
                if abs(bag[(ust, kod)] - miktar) > 1e-12:
                    self.uyari.append("Urun agacinda ayni bag farkli miktarla: %s -> %s "
                                      "(%g / %g) - ilki kullanildi" % (ust, kod, bag[(ust, kod)], miktar))
                continue
            bag[(ust, kod)] = miktar
            cocuklar[ust].append(kod)
        ebeveyn = set(cocuklar)

        # ---- girintili agaci bastan, tutarli kur (motor bunu okur) ----
        def dolas(kod, sev, miktar, iz):
            self.agac.append((sev, kod, miktar, 0))
            if kod in iz:
                self.uyari.append("Urun agacinda dongu: %s" % " -> ".join(list(iz) + [kod]))
                return
            for c in cocuklar.get(kod, []):
                dolas(c, sev + 1, bag[(kod, c)], iz + (kod,))
        for k in kokler:
            dolas(k, 0, 1.0, ())

        self.sure_catisma = []
        for sev, kod, r, _ust in ham_satir:
            aciklama = _s(kolon(r, "Malzeme Açıklaması", "Malzeme Aciklamasi",
                                "Açıklama", "Aciklama"))
            birim = _s(kolon(r, "Miktar Birimi", "Birim")) or "AD"
            miktar = _f(kolon(r, "Miktar"), 1.0) or 1.0
            im = _s(kolon(r, "İŞ MERKEZİ / KAPASİTE GRUBU", "IS MERKEZI / KAPASITE GRUBU",
                          "İŞ MERKEZİ", "IS MERKEZI"))
            if im in ("0", "0.0", "-"):
                im = ""
            pers = _f(kolon(r, "ÇALIŞAN PERSONEL SAYISI", "CALISAN PERSONEL SAYISI"), 1) or 1
            sure = _f(kolon(r, "ROTA SÜRESİ (sn)", "ROTA SURESI (sn)", "ROTA SÜRESİ"), 0.0)

            if kod in self.malzeme:
                onc = self.malzeme[kod]
                if onc["im"] != im and onc["im"] and im:
                    self.sure_catisma.append((kod, onc["im"], im))
                continue

            uretim = bool(im) or (kod in ebeveyn)
            # FIRIN adimlari adet suresiyle degil PARTI ile calisir:
            # FIRINALAMA = "20000 AD/ 20 H" -> 20.000 adetlik parti 20 saat.
            parti = None
            firin = ";" in im and "FIRIN" in im.upper()
            if firin:
                ham_p = kolon(r, "FIRINALAMA", "FIRINLAMA")
                parti = parti_coz(ham_p)
                if parti and (parti[0] <= 1 or parti[1] <= 0):
                    self.uyari.append("Firin partisi gecersiz: %s ('%s') - kapasite kisiti "
                                      "uygulanmadi" % (kod, _s(ham_p)))
                    parti = None
                    sure = 0.0
                elif parti:
                    sure = parti[1] * 3600.0 / parti[0]       # esdeger adet suresi
                elif _s(ham_p):
                    self.uyari.append("FIRINALAMA okunamadi: %s ('%s') - rota suresi "
                                      "kullanildi" % (kod, _s(ham_p)))
                else:
                    self.uyari.append("FIRINALAMA bos: %s - rota suresi (%g sn/adet) kullanildi"
                                      % (kod, sure))
            if uretim and sure <= 0 and im:
                self.uyari.append("Rota suresi bos/0: %s (%s) - kapasite kisiti uygulanmadi"
                                  % (kod, im))
            self.malzeme[kod] = dict(kod=kod, aciklama=aciklama, birim=birim,
                                     tip="URETIM" if uretim else "SATINALMA",
                                     im=im, sure=sure, pers=pers, stok=0.0,
                                     grup=self._grup_adi(im), parti=parti)
            if uretim:
                etkin = sure if sure > 0 else TANIMSIZ_SN
                for onc_no, rk, sn in self._is_merkezi(kod, im, etkin, pers):
                    self.rota.append((kod, onc_no, rk, round(sn, 6),
                                      max(1, int(round(pers or 1)))))

        # ayni kalip-makine cifti birden fazla kez eklenmis olabilir
        gor = set()
        tek = []
        for k, o, m, c in self.kalipmakine:
            if (k, m) in gor:
                continue
            gor.add((k, m))
            tek.append((k, o, m, c))
        self.kalipmakine = tek

    def _kaynak_ayarlarini_uygula(self):
        """ayarlar.json > kaynaklar: tek tek kaynagin saat/paralel/gunlerini ezer."""
        for kod, d in self.kaynak_ayar.items():
            if kod not in self.kaynak or not isinstance(d, dict):
                continue
            k = self.kaynak[kod]
            if "saat" in d:
                k["saat"] = float(d["saat"])
            if "paralel" in d:
                k["paralel"] = max(1, int(d["paralel"]))
            if "gunler" in d and len(d["gunler"]) == 7:
                k["gunler"] = [1 if x else 0 for x in d["gunler"]]

    # ----------------------------------------------------------------- #
    def _stok(self):
        toplam = defaultdict(float)
        self.depo_toplam = defaultdict(float)
        self.haric_stok = defaultdict(float)
        for r in oku(self.wb, "STOK"):
            m = _s(kolon(r, "Malzeme"))
            if not m:
                continue
            mik = _f(kolon(r, "Kullanılabilir Stok", "Kullanilabilir Stok"))
            depo = _s(kolon(r, "Depo", "Depo Kodu", "Ambar"))
            self.depo_toplam[depo] += mik
            if depo in self.haric_depolar:          # plana alinmayan depolar
                self.haric_stok[depo] += mik
                continue
            toplam[m] += mik
        self.stok_kayitsiz = 0
        for kod, m in self.malzeme.items():
            if kod in toplam:
                m["stok"] = round(toplam[kod], 6)
            else:
                self.stok_kayitsiz += 1
        self.stok_toplam = sum(self.malzeme[k]["stok"] for k in self.malzeme)

    # ----------------------------------------------------------------- #
    def _siparis(self):
        self.agacsiz = defaultdict(float)
        for r in oku(self.wb, "SİPARİŞ") or oku(self.wb, "SIPARIS"):
            m = _s(kolon(r, "Malzeme"))
            if not m:
                continue
            mik = _f(kolon(r, "Açık Miktar", "Acik Miktar"))
            tar = _d(kolon(r, "Teslim Tarihi"))
            if mik <= 0 or tar is None:
                continue
            if m not in self.malzeme:
                self.agacsiz[m] += mik
                continue
            self.talep.append(dict(
                belge=_s(kolon(r, "Sal. Bel. Num.", "Belge No")) or _s(kolon(r, "Harici Sip. No")),
                kalem=_s(kolon(r, "KalemNo", "Kalem")),
                malzeme=m, miktar=mik,
                birim=_s(kolon(r, "Miktar Br.", "Birim")) or "AD",
                tarih=tar, oncelik=0))


# ===================================================================== #
#  EXCEL YAZIMI
# ===================================================================== #
BF = PatternFill("solid", fgColor="1F3864")
BFONT = Font(color="FFFFFF", bold=True, size=10)
INCE = Side(style="thin", color="D9D9D9")
KENAR = Border(left=INCE, right=INCE, top=INCE, bottom=INCE)


def sayfa(wb, ad, bas, satirlar, genis, formatlar=None, not_metni="", renk=None):
    ws = wb.create_sheet(ad)
    if renk:
        ws.sheet_properties.tabColor = renk
    r0 = 1
    if not_metni:
        c = ws.cell(row=1, column=1, value=not_metni)
        c.font = Font(italic=True, size=9, color="7F7F7F")
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(len(bas), 1))
        r0 = 2
    for j, b in enumerate(bas, start=1):
        c = ws.cell(row=r0, column=j, value=b)
        c.fill, c.font, c.border = BF, BFONT, KENAR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, s in enumerate(satirlar, start=r0 + 1):
        for j, v in enumerate(s, start=1):
            c = ws.cell(row=i, column=j, value=v)
            c.border = KENAR
            if isinstance(v, dt.date) and not isinstance(v, dt.datetime):
                c.number_format = "DD.MM.YYYY"
            elif formatlar and bas[j - 1] in formatlar:
                c.number_format = formatlar[bas[j - 1]]
    for j, g in enumerate(genis, start=1):
        ws.column_dimensions[get_column_letter(j)].width = g
    ws.freeze_panes = ws.cell(row=r0 + 1, column=1)
    ws.auto_filter.ref = "A%d:%s%d" % (r0, get_column_letter(len(bas)), r0 + max(len(satirlar), 1))
    return ws


def yaz(dn, yol, baslangic):
    wb = Workbook()
    wb.remove(wb.active)
    plan = dict(dn.ayar.get("plan") or {})

    sayfa(wb, "Malzemeler",
          ["Malzeme Kodu", "Aciklama", "Tip", "Birim", "Kaynak", "Saniye/Adet",
           "Anlik Stok", "Emniyet Stogu", "Kapasite Grubu"],
          [[m["kod"], m["aciklama"][:70], m["tip"], m["birim"], "", "", m["stok"], 0,
            m.get("grup", "")] for m in dn.malzeme.values()],
          [24, 56, 11, 8, 12, 12, 15, 13, 30],
          {"Anlik Stok": "#,##0.###"},
          "Otomatik uretildi - veri_aktar.py. Kaynak ve Saniye/Adet bos: "
          "her malzemenin kaynaklari Rotalar sayfasinda.", "2E75B6")

    def mad(kod):
        return (dn.malzeme.get(kod, {}) or {}).get("aciklama", "") or ""

    sayfa(wb, "UrunAgaciSeviyeli",
          ["Seviye", "Malzeme", "Malzeme Aciklamasi", "Miktar (alt/1 ust)", "Fire %"],
          [[a, b, mad(b), c, d] for a, b, c, d in dn.agac],
          [10, 26, 52, 20, 10],
          {"Miktar (alt/1 ust)": "#,##0.######"},
          "%s sayfasindan aktarildi; ebeveynler 'Ust Malzeme' kolonundan (yoksa girintiden) "
          "bulunup agac tutarli girintiyle yeniden kuruldu." % (dn.agac_sayfa or ""), "2E75B6")

    sayfa(wb, "Kaynaklar",
          ["Kaynak Kodu", "Aciklama", "Tip", "Saat/Gun", "Paralel Adet",
           "Pzt", "Sal", "Car", "Per", "Cum", "Cmt", "Paz",
           "Ek Kaynak", "Ek Kaynak Katsayisi"],
          [[k["kod"], k["aciklama"][:60], k["tip"], k["saat"], k["paralel"]] + k["gunler"]
           + [k.get("ek", ""), k.get("ek_kats", "")]
           for k in dn.kaynak.values()],
          [20, 54, 10, 10, 12, 5, 5, 5, 5, 5, 5, 5, 16, 15],
          {"Saat/Gun": "#,##0.00", "Ek Kaynak Katsayisi": "#,##0.00"},
          "Kisiler %g sa/gun Pzt-Cum | makine, kalip ve firinlar %g sa/gun Pzt-Cmt. "
          "Ek Kaynak: o kaynak calisirken ayni anda tuketilen ikinci kaynak - enjeksiyon "
          "makineleri %s havuzundan cekiyor (katsayi %g = 1 operator 1 makine; bir operator "
          "N makineye bakiyorsa katsayiyi 1/N yap). Firinlar operator cekmiyor."
          % (KISI_SAAT, MAKINE_SAAT, OPERATOR_HAVUZ, OPERATOR_MAKINE), "C55A11")

    sayfa(wb, "Rotalar",
          ["Malzeme", "Malzeme Aciklamasi", "Oncelik", "Kaynak", "Saniye/Adet", "Gerekli Kisi",
           "Parti Miktari", "Parti Suresi (saat)", "Kaynak Adi", "Kapasite Grubu"],
          [[m, mad(m), o, k, sn, gk,
            (dn.malzeme.get(m, {}).get("parti") or (None, None))[0],
            (dn.malzeme.get(m, {}).get("parti") or (None, None))[1],
            dn.kaynak.get(k, {}).get("aciklama", "")[:60],
            dn.malzeme.get(m, {}).get("grup", "")]
           for m, o, k, sn, gk in dn.rota], [24, 48, 10, 20, 14, 12, 13, 14, 52, 28],
          {"Saniye/Adet": "#,##0.000", "Gerekli Kisi": "0", "Parti Miktari": "#,##0",
           "Parti Suresi (saat)": "#,##0.##"},
          "KIM NE URETEBILIR. Oncelik 1 once calisir; kapasitesi dolunca kalan ayni gun "
          "2'ye tasar. 'Gerekli Kisi' = o isi ayni anda yapmasi gereken kisi sayisi "
          "(ERP'de CALISAN PERSONEL SAYISI); saniye 1 adedin EKIP tarafindan bitirilme "
          "suresidir. Firinlarda 'Parti Miktari' adetlik parti 'Parti Suresi' saatte "
          "pisirilir (FIRINALAMA kolonu); Saniye/Adet orada sadece esdeger hizdir. 'Kaynak Adi' sadece okumak icin, program kullanmaz.", "C55A11")

    kp = getattr(dn, "kapanis", []) or []
    sayfa(wb, "Kapanislar",
          ["Kaynak", "Kaynak Adi", "Tur", "Baslangic", "Bitis", "Kalan Kapasite %", "Aciklama"],
          [[k.get("kaynak", ""), dn.kaynak.get(k.get("kaynak", ""), {}).get("aciklama", "")[:50],
            k.get("tur", "IZIN"), _d(k.get("bas")), _d(k.get("bit")) or _d(k.get("bas")),
            float(k.get("oran", 0) or 0), k.get("not", "")] for k in kp],
          [18, 46, 12, 13, 13, 18, 40],
          {"Kalan Kapasite %": "0"},
          "KAYNAK KAPATMALARI - kisi izni, makine/kalip arizasi, bakim. Secilen tarih "
          "araliginda kaynagin kapasitesi 'Kalan Kapasite %' kadar olur (0 = tamamen "
          "kapali, 50 = yarim kapasite). Bu sayfa cizelge HTML'indeki 'Kaynak kapatma' "
          "ekranindan kapanislar.json olarak kaydedilir.", "C00000")

    sayfa(wb, "KalipMakine",
          ["Kalip", "Oncelik", "Makine", "Hiz Carpani", "Kalip Hangi Parca Icin",
           "Parca Aciklamasi"],
          [[k, o, m, c, dn.kalip_parca.get(k, ""), mad(dn.kalip_parca.get(k, ""))]
           for k, o, m, c in dn.kalipmakine],
          [22, 10, 16, 13, 26, 52],
          {"Hiz Carpani": "#,##0.000"},
          "HANGI KALIP HANGI MAKINEDE CALISABILIR. Enjeksiyon is merkezi listesinden "
          "uretildi; parantezdeki gecmis kullanim adedine gore soldan saga oncelik.", "C55A11")

    # ---- ATAMALAR: tek sayfada okunur ozet ----
    kal_mak = defaultdict(list)
    for k, o, m, c in sorted(dn.kalipmakine, key=lambda x: (x[0], x[1])):
        kal_mak[k].append(m)
    alt = defaultdict(list)
    for m, o, k, sn, gk in dn.rota:
        alt[m].append((o, k, sn, gk))
    satirlar = []
    for kod, mz in dn.malzeme.items():
        if mz["tip"] != "URETIM":
            continue
        for o, k, sn, gk in sorted(alt.get(kod, [])):
            kk = dn.kaynak.get(k, {})
            satirlar.append([
                kod, mz["aciklama"][:48], mz["im"] or "(is merkezi girilmemis)",
                kk.get("tip", ""), o, k, kk.get("aciklama", "")[:46], sn, gk,
                ", ".join(kal_mak.get(k, [])) or ""])
    sayfa(wb, "Atamalar",
          ["Malzeme", "Malzeme Aciklamasi", "Operasyon / Is Merkezi", "Kaynak Tipi",
           "Oncelik", "Kaynak Kodu", "Kaynak / Kisi Adi", "Saniye/Adet", "Gerekli Kisi",
           "Onayli Makineler (oncelik sirasiyla)"],
          satirlar, [22, 44, 26, 11, 9, 18, 42, 12, 12, 40],
          {"Saniye/Adet": "#,##0.000", "Gerekli Kisi": "0"},
          "HANGI URUNU KIM YAPABILIR / HANGI MAKINEYE ATANABILIR - tek sayfada ozet. "
          "Her satir bir alternatif: Oncelik 1 ilk tercih. Kaynak Tipi KALIP ise son "
          "kolondaki makinelerde calisabilir. Bu sayfa yalnizca okumak icindir; "
          "program Rotalar ve KalipMakine sayfalarini kullanir.", "7030A0")

    sayfa(wb, "Talepler",
          ["Belge No", "Kalem", "Malzeme", "Malzeme Aciklamasi", "Miktar", "Birim",
           "Teslim Tarihi", "Oncelik"],
          [[t["belge"], t["kalem"], t["malzeme"], mad(t["malzeme"]), t["miktar"], t["birim"],
            t["tarih"], 0]
           for t in sorted(dn.talep, key=lambda x: (x["tarih"], x["belge"], x["kalem"]))],
          [16, 8, 24, 48, 14, 8, 14, 9],
          {"Miktar": "#,##0.###"},
          "SIPARIS sayfasindaki Acik Miktar kullanildi.", "548235")

    sayfa(wb, "AcikSiparisler", ["Belge", "Malzeme", "Miktar", "Tarih", "Aciklama"], [],
          [28, 24, 14, 14, 44], None,
          "Acik is emirleri / satin alma siparisleri buraya eklenebilir "
          "(pozitif = giris, negatif = rezervasyon).", "548235")

    satirlar = []
    for g, kisiler in sorted(dn.grup.items()):
        for n, k in enumerate(kisiler, start=1):
            satirlar.append([g, n, k, dn.kaynak.get(k, {}).get("aciklama", "")[:50]])
    for g, kodlar in sorted(dn.pkt.items()):
        for n, kk in enumerate(kodlar, start=1):
            rk = dn.kisi_kodu.get(kk, kk)
            satirlar.append([g, n, rk, dn.kaynak.get(rk, {}).get("aciklama", "")[:50]])
    if dn.operator_sayisi:
        for n, k in enumerate(dn.havuz_uye, start=1):
            satirlar.append([OPERATOR_HAVUZ, n, k,
                             dn.kaynak.get(k, {}).get("aciklama", "")[:50]])
    sayfa(wb, "Gruplar", ["Grup", "Sira", "Kaynak Kodu", "Kisi Adi"], satirlar,
          [34, 7, 16, 52], None,
          "Kapasite gruplari ve havuzlarin uyeleri. Cizelge HTML'i is emrinde "
          "'gercekte kim yapacak' listesini buradan uretir.", "7030A0")

    sayfa(wb, "Parametreler", ["Parametre", "Deger", "Aciklama"], [
        ["Plan Baslangic Tarihi", _d(plan.get("baslangic")) or baslangic, "Ileri planlamanin basladigi gun"],
        ["Aktarma Kurali", plan.get("aktarma", "GUNLUK"), "Alt seviyenin 1 gunluk uretimi birikince ust seviye baslar"],
        ["Aktarma Partisi", plan.get("aktarma_partisi", 0), ""],
        ["Planlama Ufku (gun)", plan.get("ufuk", 1100), "Ileri simulasyonun kac gun surecegi"],
        ["Siralama Kurali", plan.get("siralama", "EDD"), "En erken teslim tarihi once"],
        ["Tatiller", plan.get("tatiller", ""), "30.08.2026;29.10.2026 gibi"],
        ["Senaryo Analizi", plan.get("senaryo", "HAYIR"), "Buyuk veride uzun surer; EVET yapilabilir"],
        ["Satinalma Kisiti", plan.get("satinalma", "HAYIR"), "Hammadde her zaman bulunur varsayilir, acik raporlanir"],
        ["Minimum Kalip Kalis (gun)", plan.get("min_kalip_kalis", 0), "Kalip bir makineye baglandiktan sonra en az kac gun kalsin"],
        ["Gunluk Is Cesidi", plan.get("gunluk_is_cesidi", 1),
         "Bir kaynak gun icinde en fazla kac FARKLI urune gecebilir. 1 = gun icinde tek is "
         "(onerilen), 2-3 = esnek, 0 = sinirsiz. Teslim tarihi gelmis/gecmis isler bu kurali asabilir."],
        ["Is Basina Azami Kaynak", plan.get("is_basina_kaynak", 1),
         "Bir is ayni gun en fazla kac kisiye/ekibe bolunsun. 1 = isi erken bitirmek icin "
         "baskalarina bolme, tek kisi birkac gun calissin (onerilen). 0 = sinirsiz. "
         "Teslim tarihi gelmis/gecmis isler bu kurali asabilir."],
        ["Takviye Emniyet Payi (gun)", plan.get("takviye_payi", 2),
         "Takviye hesabinda birakilan pay. Is tek kisiyle teslime yetismiyorsa plan ekibi "
         "buyutur; bu pay kadar ERKEN bitirecek sekilde hesaplar. 0 = tam gununde, "
         "2 = iki is gunu once (onerilen)."],
        ["Erken Sevkiyat", plan.get("erken_sevkiyat", "HAYIR"),
         "Mal erken hazir olsa bile teslim tarihinden once sevk edilsin mi? "
         "HAYIR = stokta bekler, cikis teslim haftasinda gorunur (onerilen). "
         "EVET = hazir olur olmaz sevk edilir."],
        ["Haric Depolar", ";".join(sorted(dn.haric_depolar)) or "-",
         "Stoklari plana ALINMAYAN depo kodlari (ayarlar.json > plan.haric_depolar)."],
    ], [26, 20, 74], None, "", "7F7F7F")

    ws = wb.create_sheet("AktarimRaporu")
    ws.column_dimensions["A"].width = 130
    ws.cell(row=1, column=1, value="AKTARIM RAPORU - %s" % dt.datetime.now().strftime("%d.%m.%Y %H:%M")
            ).font = Font(bold=True, size=12, color="1F3864")
    sat = 3
    for satir in dn.rapor:
        c = ws.cell(row=sat, column=1, value=satir)
        if satir.startswith("UYARI") or satir.startswith("!"):
            c.font = Font(color="BF8F00", size=10)
        elif satir and not satir.startswith("   "):
            c.font = Font(bold=True, size=10)
        sat += 1
    ws.sheet_properties.tabColor = "BF8F00"

    wb.save(yol)


# ===================================================================== #
def main():
    girdi = sys.argv[1] if len(sys.argv) > 1 else "MRP_Girdi.xlsx"
    cikti = sys.argv[2] if len(sys.argv) > 2 else "MRP_Plan_Girdi.xlsx"
    baslangic = _d(sys.argv[3]) if len(sys.argv) > 3 else dt.date.today()

    ayar = ayarlari_oku(os.path.dirname(os.path.abspath(girdi)) or ".")
    if ayar:
        print("Ayarlar okundu  : %s (%s guncellendi)"
              % (AYAR_DOSYA, ayar.get("kaydedildi", "?")))
    print("Ham veri okunuyor:", girdi)
    dn = Donustur(girdi, ayar)
    dn.kapanis = kapanislari_oku(os.path.dirname(os.path.abspath(girdi)) or ".")
    if dn.kapanis:
        print("Kapatmalar      : %s (%d kayit)" % (KAPANIS_DOSYA, len(dn.kapanis)))

    uretim = [m for m in dn.malzeme.values() if m["tip"] == "URETIM"]
    satin = [m for m in dn.malzeme.values() if m["tip"] == "SATINALMA"]
    kal = [k for k in dn.kaynak.values() if k["tip"] == "KALIP"]
    mak = [k for k in dn.kaynak.values() if k["tip"] == "MAKINE"]
    kisi = [k for k in dn.kaynak.values() if k["tip"] == "KISI"]
    coklu = defaultdict(int)
    for m, o, k, s, gk in dn.rota:
        coklu[m] += 1

    R = []
    R.append("OZET")
    R.append("   Malzeme            : %d  (uretim %d / satinalma %d)" % (len(dn.malzeme), len(uretim), len(satin)))
    R.append("   Urun agaci satiri  : %d  (sayfa: %s)" % (len(dn.agac), dn.agac_sayfa))
    R.append("   Kaynak             : %d  (kisi %d / makine %d / kalip %d / diger %d)"
             % (len(dn.kaynak), len(kisi), len(mak), len(kal),
                len(dn.kaynak) - len(kisi) - len(mak) - len(kal)))
    R.append("   Rota satiri        : %d  (alternatifi olan malzeme: %d)"
             % (len(dn.rota), sum(1 for v in coklu.values() if v > 1)))
    R.append("   Kalip-makine bagi  : %d" % len(dn.kalipmakine))
    ekipli = sorted({(m, gk) for m, o, k, sn, gk in dn.rota if gk > 1})
    if ekipli:
        dag = {}
        for m, gk in ekipli:
            dag.setdefault(gk, []).append(m)
        R.append("   Ekip gerektiren is : %d malzeme  (%s)"
                 % (len(ekipli),
                    ", ".join("%d kisi: %d malzeme" % (gk, len(v))
                              for gk, v in sorted(dag.items()))))
    R.append("   Kaynak takvimi     : kisi %g sa, operator %g sa, makine/kalip %g sa, firin %g sa"
             % (dn.v["kisi"]["saat"], dn.v["operator"]["saat"],
                dn.v["makine"]["saat"], dn.v["firin"]["saat"]))
    if dn.kaynak_ayar:
        R.append("   ayarlar.json ile %d kaynagin kapasitesi elle ayarlandi" % len(dn.kaynak_ayar))
    if dn.operator_sayisi:
        R.append("   Operator havuzu    : %d kisi, katsayi %g, kisit %s"
                 % (dn.operator_sayisi, dn.op_katsayi, "EVET" if dn.op_kisit else "HAYIR"))
    R.append("   Talep kalemi       : %d  toplam %s adet"
             % (len(dn.talep), format(round(sum(t["miktar"] for t in dn.talep)), ",d").replace(",", ".")))
    if dn.talep:
        R.append("   Teslim araligi     : %s .. %s"
                 % (min(t["tarih"] for t in dn.talep), max(t["tarih"] for t in dn.talep)))
    R.append("   Stoklu malzeme     : %d, stok kaydi olmayan %d (sifir alindi)"
             % (len(dn.malzeme) - dn.stok_kayitsiz, dn.stok_kayitsiz))
    if getattr(dn, "depo_toplam", None):
        icinde = ", ".join("%s: %s" % (d or "(bos)", _bicim(m))
                           for d, m in sorted(dn.depo_toplam.items())
                           if d not in dn.haric_depolar)
        R.append("   Plana giren depolar: %s" % (icinde or "-"))
        if dn.haric_stok:
            R.append("   PLAN DISI depolar  : %s  (ayarlar.json > plan.haric_depolar)"
                     % ", ".join("%s: %s" % (d, _bicim(m))
                                 for d, m in sorted(dn.haric_stok.items())))
    R.append("")
    if getattr(dn, "agac_duzeltme", None):
        R.append("UYARI - urun agacinda 'Ust Malzeme' ile girinti uyusmayan satirlar (%d):"
                 % len(dn.agac_duzeltme))
        for kod, ust_k, gir, karar in dn.agac_duzeltme:
            R.append("   %-20s Ust Malzeme=%-18s girinti=%-18s -> %s"
                     % (kod, ust_k or "-", gir or "-", karar))
        R.append("")
    if dn.agacsiz:
        R.append("UYARI - siparisi olup urun agaci olmayan mamuller (plana alinmadi):")
        for m, q in sorted(dn.agacsiz.items()):
            R.append("   %-22s %12s adet" % (m, format(round(q), ",d").replace(",", ".")))
        R.append("")
    if dn.sure_catisma:
        R.append("UYARI - ayni malzeme farkli urunlerde farkli is merkezinde gorunuyor "
                 "(ilki kullanildi):")
        for k, a, b in dn.sure_catisma[:20]:
            R.append("   %-22s %s  |  %s" % (k, a, b))
        R.append("")
    gruplu = defaultdict(set)
    for kod, m in dn.malzeme.items():
        if m["tip"] == "URETIM" and m.get("grup"):
            gruplu[m["grup"]].add(kod)
    bos = [(g, len(v)) for g, v in gruplu.items()
           if g not in (OPERATOR_GRUP, "FIRIN") and not g.upper().startswith("PKT-GRUP")
           and len(dn.grup.get(g, [])) == 0]
    if bos:
        R.append("UYARI - ARA ISLEM kapasite gruplari PERSONELLISTESI'nde tanimli degil "
                 "(her biri tek kapasiteli varsayildi):")
        for g, n in sorted(bos, key=lambda x: -x[1]):
            R.append("   %-42s %d malzemede kullaniliyor" % (g, n))
        R.append("   -> PERSONELLISTESI > KAPASITE GRUPLARI sutununa bu gruplari ekle")
        R.append("")
    ara = [(g, len(dn.grup.get(g, [])), len(v)) for g, v in gruplu.items()
           if g not in (OPERATOR_GRUP, "FIRIN") and not g.upper().startswith("PKT-GRUP")
           and len(dn.grup.get(g, [])) > 0]
    if ara:
        R.append("ARA ISLEM kapasite gruplari (ENJ/MAM disi):")
        for g, k, n in sorted(ara, key=lambda x: -x[1]):
            R.append("   %-42s %2d kisi / %d malzeme" % (g, k, n))
        R.append("")
    if dn.uyari:
        gor = {}
        for u in dn.uyari:
            gor[u] = gor.get(u, 0) + 1
        R.append("UYARI - veri eksikleri (%d):" % len(gor))
        for u in sorted(gor):
            R.append("   %s%s" % (u, "  (x%d)" % gor[u] if gor[u] > 1 else ""))
    dn.rapor = R

    for s in R:
        print(s)
    yaz(dn, cikti, baslangic)
    print()
    print("Plan girdisi yazildi:", cikti)


if __name__ == "__main__":
    main()

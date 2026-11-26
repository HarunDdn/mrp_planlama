# -*- coding: utf-8 -*-
"""
AYAR TOPLAMA

Cizelge HTML'i tarayicida calistigi icin dosyayi dogrudan program klasorune
yazamayabilir: "kaydet" dedigin dosya cogu zaman Indirilenler klasorune duser
(ya da tarayici ayni addan ikinci bir kopya icin "kapanislar (1).json" yazar).
Bu yuzden plan her calistirildiginda once burasi calisir ve Indirilenler
klasorundeki DAHA YENI ayar dosyalarini program klasorune tasir.

Tasinan dosyalar:
  ayarlar.json      - program ayarlari
  atamalar.json     - is emri kisi atamalari
  kapanislar.json   - kaynak kapatmalari (izin / ariza / bakim)

Kopyalama mtime korunarak yapilir; ayni dosya ikinci kez tasinmaz.
"""

import os
import re
import shutil
import sys

ADLAR = ["ayarlar.json", "atamalar.json", "kapanislar.json"]


def indirilen_klasorleri():
    yerler = []
    ev = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    for ad in ("Downloads", "İndirilenler", "Indirilenler"):
        yerler.append(os.path.join(ev, ad))
    od = os.environ.get("OneDrive") or os.environ.get("OneDriveCommercial")
    if od:
        yerler.append(os.path.join(od, "Downloads"))
        yerler.append(os.path.join(od, "İndirilenler"))
    gor, tekil = set(), []
    for y in yerler:
        n = os.path.normcase(os.path.abspath(y))
        if n in gor or not os.path.isdir(y):
            continue
        gor.add(n)
        tekil.append(y)
    return tekil


def adaylar(klasor, ad):
    """kapanislar.json, kapanislar (1).json, kapanislar(2).json ..."""
    kok, uzanti = os.path.splitext(ad)
    desen = re.compile(r"^%s(\s*\(\d+\))?%s$" % (re.escape(kok), re.escape(uzanti)),
                       re.IGNORECASE)
    try:
        dosyalar = os.listdir(klasor)
    except OSError:
        return []
    return [os.path.join(klasor, d) for d in dosyalar if desen.match(d)]


def topla(hedef_klasor):
    tasinan = []
    kaynaklar = indirilen_klasorleri()
    for ad in ADLAR:
        hedef = os.path.join(hedef_klasor, ad)
        h_mtime = os.path.getmtime(hedef) if os.path.exists(hedef) else 0.0
        en_yeni, en_yeni_mtime = None, h_mtime
        for klasor in kaynaklar:
            for yol in adaylar(klasor, ad):
                try:
                    m = os.path.getmtime(yol)
                except OSError:
                    continue
                if m > en_yeni_mtime + 1:          # 1 sn tolerans
                    en_yeni, en_yeni_mtime = yol, m
        if en_yeni:
            try:
                shutil.copy2(en_yeni, hedef)       # mtime korunur
                tasinan.append((ad, en_yeni))
            except OSError as e:
                print("UYARI: %s tasinamadi (%s)" % (ad, e))
    return tasinan


def main():
    hedef = os.path.dirname(os.path.abspath(sys.argv[0])) or "."
    if len(sys.argv) > 1:
        hedef = sys.argv[1]
    tasinan = topla(hedef)
    if tasinan:
        for ad, nereden in tasinan:
            print("  %-18s <- %s" % (ad, nereden))
    else:
        print("  Indirilenler klasorunde daha yeni ayar dosyasi yok.")


if __name__ == "__main__":
    main()

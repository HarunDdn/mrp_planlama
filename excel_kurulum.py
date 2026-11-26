# -*- coding: utf-8 -*-
"""
MRP_Girdi calisma kitabina CANIAS guncelleme butonlarini kurar / sorgulari yeniler.

    python excel_kurulum.py                  # MRP_Girdi.xlsm (yoksa MRP_Girdi.xlsx)
    python excel_kurulum.py baska.xlsx

Yaptiklari (Excel kapaliyken, arka planda Excel ile):
  1. Dosyanin kopyasini yedek/ klasorune alir.
  2. STOK ve SİPARİŞ sekmelerinin en ustune bir satir ekler ve oraya
     "CANIAS'tan Güncelle" butonunu koyar (zaten varsa dokunmaz).
  3. Gizli "CANIAS" sayfasina baglanti bilgisini (canias_baglanti.json) ve
     sorgulari (stok_sorgu.sql, siparis_sorgu.sql) yazar. Sorgu dosyalarini
     degistirdiginizde bu scripti tekrar calistirmaniz yeterli.
  4. Dosyayi makro icerebilen .xlsm olarak kaydeder.

Makro kodu (canias_guncelle.bas) Excel guvenlik ayari geregi elle bir kez ice
aktarilir: Alt+F11 > File > Import File... > canias_guncelle.bas > Kaydet.
"""

import datetime as dt
import json
import os
import shutil
import sys

KLASOR = os.path.dirname(os.path.abspath(__file__))
SAYFALAR = [("STOK", "stok_sorgu.sql"), ("SIPARIS", "siparis_sorgu.sql")]
BUTON = "btnCanias"
MAKRO = "CaniasGuncelle"


def _norm(s):
    s = str(s or "").strip()
    for a, b in zip("ıİğĞüÜşŞöÖçÇ", "iIgGuUsSoOcC"):
        s = s.replace(a, b)
    return s.upper()


def _bgr(r, g, b):
    return r + g * 256 + b * 65536


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if args:
        girdi = os.path.abspath(args[0])
    else:
        xlsm = os.path.join(KLASOR, "MRP_Girdi.xlsm")
        girdi = xlsm if os.path.exists(xlsm) else os.path.join(KLASOR, "MRP_Girdi.xlsx")
    if not os.path.exists(girdi):
        raise SystemExit("[HATA] %s bulunamadi." % girdi)
    kilit = os.path.join(os.path.dirname(girdi), "~$" + os.path.basename(girdi))
    if os.path.exists(kilit):
        raise SystemExit("[HATA] %s Excel'de acik. Kapatip tekrar calistirin." % os.path.basename(girdi))

    with open(os.path.join(KLASOR, "canias_baglanti.json"), encoding="utf-8") as f:
        bag = json.load(f)
    sorgular = {}
    for ad, dosya in SAYFALAR:
        with open(os.path.join(KLASOR, dosya), encoding="utf-8") as f:
            sorgular[ad] = f.read().strip()
            if len(sorgular[ad]) > 32000:
                raise SystemExit("[HATA] %s Excel hucresine sigmayacak kadar uzun." % dosya)

    yedek_k = os.path.join(os.path.dirname(girdi), "yedek")
    os.makedirs(yedek_k, exist_ok=True)
    ad, uz = os.path.splitext(os.path.basename(girdi))
    yedek = os.path.join(yedek_k, "%s_%s%s" % (ad, dt.datetime.now().strftime("%Y%m%d_%H%M%S"), uz))
    shutil.copy2(girdi, yedek)
    print("Yedek alindi :", yedek)

    # Excel, OneDrive klasorundeki dosyaya dogrudan "farkli kaydet" yaparken cokebiliyor;
    # islem yerel gecici klasorde yapilip sonuc geri kopyalanir.
    import tempfile
    import win32com.client as win32
    gecici = tempfile.mkdtemp(prefix="mrp_kurulum_")
    yerel = os.path.join(gecici, os.path.basename(girdi))
    shutil.copy2(girdi, yerel)
    yerel_hedef = os.path.splitext(yerel)[0] + ".xlsm"
    hedef = os.path.splitext(girdi)[0] + ".xlsm"
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    try:
        wb = xl.Workbooks.Open(yerel, UpdateLinks=0)
        adlar = {_norm(ws.Name): ws.Name for ws in wb.Worksheets}

        for anahtar, _ in SAYFALAR:
            gercek = adlar.get(anahtar)
            if not gercek:
                raise SystemExit("[HATA] Calisma kitabinda '%s' sekmesi yok." % anahtar)
            ws = wb.Worksheets(gercek)
            var = any(s.Name == BUTON for s in ws.Shapes)
            if not var:
                ws.Rows(1).Insert()
                ws.Rows(1).RowHeight = 30
                shp = ws.Shapes.AddShape(5, 4, 3, 175, 24)        # yuvarlak dikdortgen
                shp.Name = BUTON
                shp.Fill.ForeColor.RGB = _bgr(46, 117, 182)
                shp.Line.Visible = False
                tf = shp.TextFrame2
                tf.TextRange.Text = "⟳  CANIAS'tan Güncelle"
                tf.TextRange.Font.Size = 10
                tf.TextRange.Font.Bold = True
                tf.TextRange.Font.Fill.ForeColor.RGB = _bgr(255, 255, 255)
                tf.VerticalAnchor = 3                                # orta
                tf.TextRange.ParagraphFormat.Alignment = 2           # ortala
                shp.Placement = 3                                    # hucrelerle tasinmaz
                shp.OnAction = MAKRO
                # butonun sagindaki ilk sutuna aciklama
                c = 1
                while ws.Cells(1, c).Left < shp.Left + shp.Width + 6 and c < 50:
                    c += 1
                ws.Cells(1, c).Value = "Butona basın: güncel veri CANIAS'tan çekilir"
                ws.Cells(1, c).Font.Italic = True
                ws.Cells(1, c).Font.Color = _bgr(89, 89, 89)
                print("%-8s : buton eklendi, veri 2. satirdan basliyor" % gercek)
            else:
                print("%-8s : buton zaten var" % gercek)

        try:
            ayar = wb.Worksheets("CANIAS")
        except Exception:
            ayar = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ayar.Name = "CANIAS"
        ayar.Cells.Clear()
        ayar.Columns(2).NumberFormat = "@"                          # '--' ile baslayan sorgu formul sanilmasin
        satirlar = [("Sunucu", bag.get("sunucu", "")), ("Veritabani", bag.get("veritabani", "")),
                    ("Kullanici", bag.get("kullanici", "")),
                    ("Surucu", bag.get("surucu") or "ODBC Driver 17 for SQL Server")]
        for i, (k, v) in enumerate(satirlar, start=1):
            ayar.Cells(i, 1).Value = k
            ayar.Cells(i, 2).Value = v
        ayar.Cells(6, 1).Value = "Sayfa"
        ayar.Cells(6, 2).Value = "Sorgu (python excel_kurulum.py ile .sql dosyalarindan yuklenir)"
        for i, (anahtar, _) in enumerate(SAYFALAR, start=7):
            ayar.Cells(i, 1).Value = adlar[anahtar]
            ayar.Cells(i, 2).Value = sorgular[anahtar]
        ayar.Visible = 0                                            # xlSheetHidden
        print("CANIAS   : baglanti + %d sorgu yazildi (gizli sayfa)" % len(SAYFALAR))

        if yerel.lower().endswith(".xlsm"):
            wb.Save()
        else:
            wb.SaveAs(yerel_hedef, FileFormat=52)
        wb.Close(False)
    finally:
        xl.Quit()

    shutil.copyfile(yerel_hedef, hedef)
    shutil.rmtree(gecici, ignore_errors=True)
    if not girdi.lower().endswith(".xlsm"):
        os.remove(girdi)                                            # kopyasi yedek/ klasorunde
        print("Kaydedildi  : %s  (eski .xlsx yedekte)" % hedef)
    else:
        print("Kaydedildi  : %s" % hedef)
    print()
    print("Son adim (bir kez): dosyayi Excel'de acin > Alt+F11 > File > Import File...")
    print("  > canias_guncelle.bas > pencereyi kapatip Kaydet.")


if __name__ == "__main__":
    main()

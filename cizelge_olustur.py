# -*- coding: utf-8 -*-
"""
HAFTALIK CIZELGE GORSELI
========================

MRP_Sonuc.xlsx'i okur, tarayicida acilan tek dosyalik haftalik cizelge uretir:
hangi kaynak hangi gun hangi uretimde kullanilmis.

    python cizelge_olustur.py MRP_Sonuc.xlsx MRP_Cizelge.html
"""

import datetime as dt
import json
import os
import sys
from collections import OrderedDict, defaultdict

from openpyxl import load_workbook


def _js(yol):
    if not os.path.exists(yol):
        return {}
    try:
        with open(yol, encoding="utf-8") as f:
            return json.load(f) or {}
    except (OSError, ValueError):
        return {}


def olustur(girdi="MRP_Sonuc.xlsx", cikti="MRP_Cizelge.html",
            plan_girdi="MRP_Plan_Girdi.xlsx"):
    wb = load_workbook(girdi, data_only=True)
    klasor = os.path.dirname(os.path.abspath(girdi)) or "."

    # --- plan girdisi: kaynak takvimi, rotalar, gruplar, parametreler
    kay_ayar, rota, gruplar, param = {}, {}, {}, {}
    pg = None
    if os.path.exists(os.path.join(klasor, os.path.basename(plan_girdi))):
        pg = load_workbook(os.path.join(klasor, os.path.basename(plan_girdi)), data_only=True)
        if "Kaynaklar" in pg.sheetnames:
            w = pg["Kaynaklar"]
            hb = [c.value for c in next(w.iter_rows(min_row=2, max_row=2))]
            jb = {b: i for i, b in enumerate(hb)}
            for r in w.iter_rows(min_row=3, values_only=True):
                if not r[0]:
                    continue
                kay_ayar[r[0]] = {
                    "ad": r[jb.get("Aciklama", 1)] or "", "tip": r[jb.get("Tip", 2)] or "",
                    "saat": r[jb.get("Saat/Gun", 3)] or 0,
                    "paralel": r[jb.get("Paralel Adet", 4)] or 1,
                    "gunler": [1 if r[jb.get(g, 5 + i)] else 0
                               for i, g in enumerate(("Pzt", "Sal", "Car", "Per",
                                                      "Cum", "Cmt", "Paz"))],
                    "ek": r[jb.get("Ek Kaynak", 12)] or "",
                    "ekKats": r[jb.get("Ek Kaynak Katsayisi", 13)] or ""}
        if "Rotalar" in pg.sheetnames:
            w = pg["Rotalar"]
            hr = [c.value for c in next(w.iter_rows(min_row=2, max_row=2))]
            jr = {b: i for i, b in enumerate(hr) if b}
            i_m, i_o, i_k = jr.get("Malzeme", 0), jr.get("Oncelik", 1), jr.get("Kaynak", 2)
            for r in w.iter_rows(min_row=3, values_only=True):
                if r[i_m] and r[i_k]:
                    rota.setdefault(r[i_m], []).append([r[i_o] or 0, r[i_k]])
        if "Gruplar" in pg.sheetnames:
            w = pg["Gruplar"]
            for r in w.iter_rows(min_row=3, values_only=True):
                if r[0] and r[2]:
                    gruplar.setdefault(r[0], []).append(r[2])
        if "Parametreler" in pg.sheetnames:
            for r in pg["Parametreler"].iter_rows(min_row=2, values_only=True):
                if r and r[0]:
                    v = r[1]
                    param[str(r[0]).strip()] = (v.date().isoformat()
                                                if isinstance(v, dt.datetime) else v)

    kayitli = _js(os.path.join(klasor, "ayarlar.json"))
    atamalar = _js(os.path.join(klasor, "atamalar.json"))
    kapanis_d = _js(os.path.join(klasor, "kapanislar.json")) or {}
    kapanis = kapanis_d.get("kapanislar") if isinstance(kapanis_d, dict) else kapanis_d
    kapanis = [k for k in (kapanis or []) if isinstance(k, dict) and k.get("kaynak")]
    kapanis_kaynagi = "kapanislar.json" if kapanis else ""
    # JSON bulunamadiysa plana FIILEN giren kapatmalari plan girdisinden al:
    # boylece dosya yanlis klasore kaydedilmis olsa bile ekran bos kalmaz.
    if not kapanis and pg is not None and "Kapanislar" in pg.sheetnames:
        wk = pg["Kapanislar"]
        hk = None
        for r in wk.iter_rows(min_row=1, max_row=4, values_only=True):
            if r and r[0] == "Kaynak":
                hk = {b: i for i, b in enumerate(r)}
                break
        if hk:
            for r in wk.iter_rows(min_row=3, values_only=True):
                if not r[hk["Kaynak"]] or r[hk["Kaynak"]] == "Kaynak":
                    continue
                g = lambda a: r[hk[a]] if a in hk else None
                t1, t2 = g("Baslangic"), g("Bitis")
                kapanis.append({
                    "kaynak": str(g("Kaynak")),
                    "bas": t1.date().isoformat() if isinstance(t1, dt.datetime) else str(t1 or ""),
                    "bit": t2.date().isoformat() if isinstance(t2, dt.datetime) else str(t2 or ""),
                    "tur": str(g("Tur") or "IZIN"),
                    "oran": float(g("Kalan Kapasite %") or 0),
                    "not": str(g("Aciklama") or "")})
        if kapanis:
            kapanis_kaynagi = "plan girdisi (Kapanislar sayfasi)"
    ws = wb["UretimPlani"]
    bas = [c.value for c in next(ws.iter_rows(max_row=1))]
    ix = {b: i for i, b in enumerate(bas)}

    mlz = OrderedDict()      # malzeme -> aciklama
    kyn = OrderedDict()      # kaynak -> [ad, tip, saat/gun, gunler]
    kayit = []
    yuk_tab = {}             # kaynak -> gun -> yuklenen saat (motorun kendi rakami)

    # Kaynak takvimi ve kapasitesi KapasiteYuku sayfasindan: TUM kaynaklari
    # icerir (kisi, kalip, makine, firin, operator havuzu) - UretimPlani'nda
    # makineler ayri satir degil, kalibin yanindaki alan oldugu icin sart.
    if "KapasiteYuku" in wb.sheetnames:
        wk = wb["KapasiteYuku"]
        bk = [c.value for c in next(wk.iter_rows(max_row=1))]
        jk = {b: i for i, b in enumerate(bk)}
        for r in wk.iter_rows(min_row=2, values_only=True):
            k = r[jk["Kaynak"]]
            g = r[jk["Tarih"]]
            if not k or g is None:
                continue
            g = g.date() if isinstance(g, dt.datetime) else g
            kp = r[jk["Kapasite (saat)"]] or 0
            yk = r[jk["Yuk (saat)"]] or 0
            if yk > 0:
                yuk_tab.setdefault(k, {})[g.isoformat()] = round(yk, 3)
            if k not in kyn:
                kyn[k] = [r[jk.get("Kaynak / Kisi Adi", 3)] or "",
                          r[jk.get("Tip", 4)] or "", 0.0, [0] * 7]
            if kp > 0:
                kyn[k][2] = max(kyn[k][2], round(kp, 2))
                kyn[k][3][g.weekday()] = 1

    kix = {k: i for i, k in enumerate(kyn)}

    def kaynak_ix(kod, ad="", tip=""):
        if kod not in kix:
            kyn[kod] = [ad, tip, 0.0, [1, 1, 1, 1, 1, 0, 0]]
            kix[kod] = len(kix)
        return kix[kod]

    for r in ws.iter_rows(min_row=2, values_only=True):
        g = r[ix["Tarih"]]
        if g is None:
            continue
        g = g.date() if isinstance(g, dt.datetime) else g
        m = r[ix["Malzeme"]]
        if m not in mlz:
            mlz[m] = (r[ix["Aciklama"]] or "")
        ki = kaynak_ix(r[ix["Kaynak"]], r[ix["Kaynak / Kisi Adi"]] or "",
                       r[ix["Kaynak Tipi"]] or "")
        mak = r[ix["Makine"]] or ""
        mi = kaynak_ix(mak, "Enjeksiyon makinesi", "MAKINE") if mak else -1
        kayit.append([g.isoformat(), list(mlz).index(m), ki, mi,
                      round(r[ix["Miktar"]] or 0, 1),
                      round(r[ix["Yuk (saat)"]] or 0, 3),
                      (r[ix["Kapasite Grubu"]] or "") if "Kapasite Grubu" in ix else ""])

    # is emirleri (toplulastirilmis): ayni urun + ayni kaynak + ayni hafta
    emirler = []
    if "IsEmirleri" in wb.sheetnames:
        we = wb["IsEmirleri"]
        be = [c.value for c in next(we.iter_rows(max_row=1))]
        je = {b: i for i, b in enumerate(be)}
        for r in we.iter_rows(min_row=2, values_only=True):
            if not r[je["Is Emri"]] or not r[je["Urun"]]:
                continue
            u = r[je["Urun"]]
            if u not in mlz:
                mlz[u] = r[je["Aciklama"]] or ""
            il, sn = r[je["Baslangic"]], r[je["Bitis"]]
            il = il.date() if isinstance(il, dt.datetime) else il
            sn = sn.date() if isinstance(sn, dt.datetime) else sn
            mak = r[je["Makine"]] or ""
            emirler.append([
                r[je["Is Emri"]], list(mlz).index(u),
                kaynak_ix(r[je["Kaynak"]], r[je["Kaynak / Kisi Adi"]] or "",
                          r[je["Tip"]] or ""),
                kaynak_ix(mak, "Enjeksiyon makinesi", "MAKINE") if mak else -1,
                il.isoformat(), sn.isoformat(), r[je["Gun"]] or 1,
                round(r[je["Miktar"]] or 0, 1), round(r[je["Yuk (saat)"]] or 0, 2),
                r[je["Tetikleyen siparis"]] or "",
                (r[je["Kisi Grubu"]] if "Kisi Grubu" in je else ""),
                (r[je["Kapasite Grubu"]] if "Kapasite Grubu" in je else ""),
                int(r[je["Gerekli Kisi"]] or 1) if "Gerekli Kisi" in je else 1,
                [kaynak_ix(k.strip(), "", "KISI")
                 for k in str(r[je["Ekip (diger kisiler)"]] or "").split(",")
                 if k.strip()] if "Ekip (diger kisiler)" in je else []])

    # ozet gostergeler
    ozet = {}
    try:
        wso = wb["Ozet"]
        for r in wso.iter_rows(values_only=True):
            if r and r[0] in ("Plan baslangic tarihi", "Toplam talep miktari",
                              "Zamaninda karsilanan kalem", "Geciken kalem",
                              "Son uretim gunu (ileri plan)", "Darbogaz kaynak"):
                v = r[1]
                ozet[r[0]] = (v.date().isoformat()
                              if isinstance(v, dt.datetime) else v)
    except KeyError:
        pass

    # geciken siparisler (hafta uyarisi icin)
    gec = []
    try:
        wsg = wb["SiparisOzeti"]
        b2 = [c.value for c in next(wsg.iter_rows(max_row=1))]
        j = {b: i for i, b in enumerate(b2)}
        for r in wsg.iter_rows(min_row=2, values_only=True):
            d = r[j["Gecikme (gun)"]]
            if isinstance(d, (int, float)) and d > 0:
                bitis = r[j["Ileri Plan: Bitis"]]
                gec.append({
                    "belge": r[j["Belge"]], "kalem": r[j["Kalem"]],
                    "malzeme": r[j["Malzeme"]], "miktar": r[j["Miktar"]],
                    "teslim": (r[j["Teslim Tarihi"]].date().isoformat()
                               if isinstance(r[j["Teslim Tarihi"]], dt.datetime) else ""),
                    "bitis": (bitis.date().isoformat()
                              if isinstance(bitis, dt.datetime) else ""),
                    "gun": int(d)})
    except KeyError:
        pass

    # --- AKIS: haftalik malzeme akisi (urun agaci yapisina gore)
    akis = {"hafta": {}, "seviye": {}}
    try:
        from akis_veri import akis_verisi
        mal_ad, mal_tip = {}, {}
        if pg is not None and "Malzemeler" in pg.sheetnames:
            wm = pg["Malzemeler"]
            for r in wm.iter_rows(min_row=3, values_only=True):
                if r[0]:
                    mal_ad[r[0]] = r[1] or ""
                    mal_tip[r[0]] = r[2] or ""
        akis = akis_verisi(wb, pg, mlz, emirler, mal_ad)
        akis["tip"] = {str(i): ("S" if mal_tip.get(k) == "SATINALMA" else "U")
                       for i, k in enumerate(mlz)}
    except Exception as hata:                       # akis olmasa da cizelge calissin
        akis = {"hafta": {}, "seviye": {}, "tip": {}, "hata": str(hata)}

    veri = {
        "malzemeler": [[k, v] for k, v in mlz.items()],
        "akis": akis,
        "kaynaklar": [[k, v[0], v[1], v[2], v[3]] for k, v in kyn.items()],
        "kayitlar": kayit,
        "emirler": emirler,
        "ozet": ozet,
        "yuk": yuk_tab,
        "kaynakAyar": kay_ayar,
        "rota": rota,
        "gruplar": gruplar,
        "param": param,
        "kayitliAyar": kayitli,
        "atamalar": atamalar,
        "kapanis": kapanis,
        "kapanisKaynagi": kapanis_kaynagi,
        "geciken": gec,
        "uretildi": dt.datetime.now().strftime("%d.%m.%Y %H:%M"),
    }

    html = SABLON.replace("/*__VERI__*/", json.dumps(veri, ensure_ascii=False))
    with open(cikti, "w", encoding="utf-8") as f:
        f.write(html)
    return cikti, len(kayit), len(kyn), len(mlz)


SABLON = r"""<!doctype html>
<html lang="tr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Haftalık Üretim Çizelgesi</title>
<style>
:root{
  color-scheme: light;
  --plane:#f9f9f7; --surface:#fcfcfb;
  --ink:#0b0b0b; --ink2:#52514e; --muted:#898781;
  --grid:#e1e0d9; --axis:#c3c2b7; --ring:rgba(11,11,11,.10);
  --s100:#cde2fb; --s150:#b7d3f6; --s200:#9ec5f4; --s250:#86b6ef;
  --s300:#6da7ec; --s350:#5598e7; --s400:#3987e5; --s450:#2a78d6;
  --s500:#256abf; --s550:#1c5cab; --s600:#184f95; --s650:#104281; --s700:#0d366b;
  --good:#0ca30c; --warning:#fab219; --serious:#ec835a; --critical:#d03b3b;
  --accent:#2a78d6;
  /* rampa adimlarinin uzerine yazilacak murekkep - acik dolgularda koyu, koyuda beyaz */
  --oi1:#0b0b0b; --oi2:#0b0b0b; --oi3:#0b0b0b; --oi4:#fff; --oi5:#fff; --oi6:#fff; --oi7:#fff;
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    color-scheme: dark;
    --plane:#0d0d0d; --surface:#1a1a19;
    --ink:#f0efec; --ink2:#c3c2b7; --muted:#898781;
    --grid:#2c2c2a; --axis:#383835; --ring:rgba(255,255,255,.10);
    --s100:#0d366b; --s150:#104281; --s200:#184f95; --s250:#1c5cab;
    --s300:#256abf; --s350:#2a78d6; --s400:#3987e5; --s450:#5598e7;
    --s500:#6da7ec; --s550:#86b6ef; --s600:#9ec5f4; --s650:#b7d3f6; --s700:#cde2fb;
    --accent:#3987e5;
    --oi1:#fff; --oi2:#fff; --oi3:#fff; --oi4:#fff; --oi5:#0b0b0b; --oi6:#0b0b0b; --oi7:#0b0b0b;
  }
}
:root[data-theme="dark"]{
  color-scheme: dark;
  --plane:#0d0d0d; --surface:#1a1a19;
  --ink:#f0efec; --ink2:#c3c2b7; --muted:#898781;
  --grid:#2c2c2a; --axis:#383835; --ring:rgba(255,255,255,.10);
  --s100:#0d366b; --s150:#104281; --s200:#184f95; --s250:#1c5cab;
  --s300:#256abf; --s350:#2a78d6; --s400:#3987e5; --s450:#5598e7;
  --s500:#6da7ec; --s550:#86b6ef; --s600:#9ec5f4; --s650:#b7d3f6; --s700:#cde2fb;
  --accent:#3987e5;
  --oi1:#fff; --oi2:#fff; --oi3:#fff; --oi4:#fff; --oi5:#0b0b0b; --oi6:#0b0b0b; --oi7:#0b0b0b;
}
*{box-sizing:border-box}
body{margin:0;background:var(--plane);color:var(--ink);
  font:14px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif;}
.wrap{max-width:1500px;margin:0 auto;padding:24px 16px 72px}
h1{font-size:21px;margin:0 0 2px;letter-spacing:-.01em}
.sub{color:var(--ink2);font-size:13px;margin:0 0 20px}

/* ---- filtre satiri ---- */
.bar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;
  background:var(--surface);border:1px solid var(--ring);border-radius:10px;
  padding:10px 12px;margin-bottom:16px;position:sticky;top:0;z-index:20}
button,select,input{font:inherit;color:var(--ink);background:var(--surface);
  border:1px solid var(--axis);border-radius:7px;padding:6px 10px}
button{cursor:pointer}
button:hover{background:var(--grid)}
button.pri{background:var(--accent);border-color:var(--accent);color:#fff;font-weight:600}
button.pri:hover{filter:brightness(1.08)}
button[aria-pressed="true"]{background:var(--accent);border-color:var(--accent);color:#fff}
.spacer{flex:1}
.hint{color:var(--muted);font-size:12px}

/* ---- stat tiles ---- */
.tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));
  gap:10px;margin-bottom:18px}
.tile{background:var(--surface);border:1px solid var(--ring);border-radius:10px;padding:12px 14px}
.tile .lbl{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.05em}
.tile .val{font-size:25px;font-weight:650;margin-top:3px;letter-spacing:-.02em}
.tile .foot{font-size:12px;color:var(--ink2);margin-top:2px}

/* ---- isi haritasi ---- */
.card{background:var(--surface);border:1px solid var(--ring);border-radius:10px;
  padding:14px;margin-bottom:18px;overflow-x:auto}
.card h2{font-size:14px;margin:0 0 3px;letter-spacing:.01em}
.card .note{font-size:12px;color:var(--ink2);margin:0 0 12px}
table.hm{border-collapse:separate;border-spacing:2px;width:100%}
table.hm th{font-weight:600;font-size:12px;color:var(--ink2);text-align:center;
  padding:4px 6px;white-space:nowrap}
table.hm th.row{text-align:left;position:sticky;left:0;background:var(--surface);
  z-index:2;min-width:230px;max-width:230px}
table.hm td{padding:0}
.grp td,.grp th{background:transparent!important;padding-top:12px!important}
.grp th{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.06em;
  font-weight:700}
.rk{display:block;font-weight:600;font-size:12.5px;white-space:nowrap;overflow:hidden;
  text-overflow:ellipsis}
.rn{display:block;font-size:11px;color:var(--muted);font-weight:400;white-space:nowrap;
  overflow:hidden;text-overflow:ellipsis}
.cell{width:100%;height:40px;border:1px solid var(--ring);border-radius:4px;background:var(--grid);
  display:flex;flex-direction:column;align-items:center;justify-content:center;
  cursor:pointer;font:inherit;color:var(--ink);padding:0;gap:0;position:relative}
.cell:focus-visible{outline:2px solid var(--accent);outline-offset:1px}
.cell.empty{background:transparent;border:1px dashed var(--grid);cursor:default}
.cell.off{background:repeating-linear-gradient(135deg,transparent,transparent 4px,
  var(--grid) 4px,var(--grid) 5px);cursor:default}
.cell .h{font-size:12.5px;font-weight:650;font-variant-numeric:tabular-nums;line-height:1.15}
.cell .p{font-size:10px;opacity:.9;font-variant-numeric:tabular-nums;line-height:1.15}
.cell.sel{outline:2px solid var(--ink);outline-offset:1px}
td.tot .cell{background:transparent;border:1px solid var(--grid);cursor:default}

/* ---- tooltip ---- */
#tip{position:fixed;z-index:60;pointer-events:none;max-width:340px;
  background:var(--surface);color:var(--ink);border:1px solid var(--axis);
  border-radius:8px;padding:9px 11px;font-size:12.5px;
  box-shadow:0 6px 24px rgba(0,0,0,.18);opacity:0;transition:opacity .08s}
#tip.on{opacity:1}
#tip b{font-size:13px}
#tip .r{display:flex;justify-content:space-between;gap:14px;margin-top:3px}
#tip .m{color:var(--muted)}
#tip hr{border:0;border-top:1px solid var(--grid);margin:7px 0}

/* ---- detay ---- */
.det{background:var(--surface);border:1px solid var(--ring);border-radius:10px;padding:14px;margin-bottom:18px}
table.dt{border-collapse:collapse;width:100%;font-size:13px}
table.dt th{text-align:left;font-size:11px;color:var(--muted);text-transform:uppercase;
  letter-spacing:.05em;border-bottom:1px solid var(--axis);padding:6px 8px;white-space:nowrap}
table.dt td{border-bottom:1px solid var(--grid);padding:6px 8px;vertical-align:top}
table.dt td.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
table.dt td.nw{white-space:nowrap}
table.dt td.nw b{font-variant-numeric:tabular-nums}
table.dt tr:hover td{background:var(--grid)}
.pill{display:inline-block;font-size:10.5px;padding:1px 7px;border-radius:999px;
  border:1px solid var(--axis);color:var(--ink2);white-space:nowrap}
.pill.warn{border-color:var(--s500);color:var(--s700);font-weight:600}
.tabs.ana{margin:18px 0 16px;gap:2px;border-bottom:1px solid var(--axis)}
.tabs.ana button{padding:10px 16px;font-size:14px;border-bottom-width:2px}
.tabs.ana button .sy{font-size:11.5px;color:var(--muted);margin-left:6px;font-weight:600}
.tabs.ana button[aria-selected="true"] .sy{color:var(--accent)}
/* ---- kaynak cizelgesi (hangi kaynakta hangi is emri) ---- */
.kc-sar{max-height:72vh;overflow:auto;border:1px solid var(--grid);border-radius:10px}
.kc{display:grid;grid-template-columns:238px 1fr;background:var(--surface);min-width:860px}
.kc-ad{position:sticky;left:0;background:var(--surface);z-index:1}
.kc-bas{position:sticky;left:0;z-index:3}
.kc-bas,.kc-gunler{background:var(--plane);border-bottom:1px solid var(--axis);
  padding:8px 10px;font-size:11px;color:var(--muted);text-transform:uppercase;
  letter-spacing:.04em;font-weight:600;position:sticky;top:0;z-index:2}
.kc-gunler{display:grid;grid-template-columns:repeat(7,1fr);padding:0}
.kc-gunler span{padding:8px 8px;border-left:1px solid var(--grid)}
.kc-gunler span.tatil{background:var(--plane);color:var(--axis)}
.kc-ad{padding:8px 10px;border-bottom:1px solid var(--grid);border-right:1px solid var(--grid);
  font-size:12px;display:flex;flex-direction:column;gap:2px;justify-content:center}
.kc-ad b{font-size:12.5px}
.kc-ad span{color:var(--muted);font-size:10.5px}
.kc-trk{position:relative;border-bottom:1px solid var(--grid);
  background:repeating-linear-gradient(to right,transparent 0,transparent calc(100%/7 - 1px),
    var(--grid) calc(100%/7 - 1px),var(--grid) calc(100%/7))}
.kc-bar{position:absolute;height:30px;border-radius:7px;padding:0 8px;overflow:hidden;
  display:flex;align-items:center;gap:7px;font-size:11.5px;white-space:nowrap;cursor:pointer;
  background:color-mix(in srgb,var(--accent) 16%,var(--surface));
  border:1px solid var(--accent);color:var(--ink)}
.kc-bar.kalip{background:color-mix(in srgb,var(--s300) 24%,var(--surface));border-color:var(--s500)}
.kc-bar.makine{background:color-mix(in srgb,var(--s200) 26%,var(--surface));border-color:var(--s400)}
.kc-bar.ekip{border-style:dashed}
.kc-bar.gec{border-left:5px solid var(--critical)}
.kc-bar b{font-weight:700;flex:0 0 auto}
.kc-bar em{font-style:normal;color:var(--ink2);flex:1 1 auto;min-width:0;
  overflow:hidden;text-overflow:ellipsis}
.kc-bar i{font-style:normal;color:var(--ink2);flex:0 0 auto}
.kc-bar i.mk{font-weight:600;color:var(--ink)}
.kc-bar i.no{opacity:.55;flex:0 1 auto;overflow:hidden;text-overflow:ellipsis}
.kc-bar.sec{outline:2px solid var(--accent);outline-offset:1px;z-index:4}
/* Dar cubukta yazi sigmiyorsa: uzerine gelince cubuk gercek genisligine acilir */
.kc-bar:hover{width:max-content !important;max-width:none;z-index:6;
  box-shadow:0 6px 18px rgba(0,0,0,.22)}
.kc-bos{padding:10px;color:var(--muted);font-size:12.5px}
.kapform{display:flex;gap:10px;flex-wrap:wrap;align-items:flex-end;margin:12px 0 16px;
  padding:12px;border:1px solid var(--grid);border-radius:10px;background:var(--plane)}
.kapform label{display:flex;flex-direction:column;gap:4px;font-size:11.5px;color:var(--ink2)}
.kapform label.gen{flex:1 1 220px}
.kapform input,.kapform select{padding:5px 7px;font-size:13px}
.kapform input[type=text]{width:100%}
.cell.kapali{background:repeating-linear-gradient(45deg,transparent 0 4px,
  color-mix(in srgb,var(--critical) 26%,transparent) 4px 8px);
  color:var(--critical);font-weight:700;display:flex;align-items:center;justify-content:center}
.akis-bar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:10px 0 6px}
.akis-bar input[type=search]{width:190px}
.akis-kutu{overflow:auto;max-height:78vh;border:1px solid var(--grid);
  border-radius:10px;background:var(--plane)}
.akis-kutu svg{display:block}
.ak-n{cursor:pointer}
.ak-n rect.gov{fill:var(--surface);stroke:var(--axis);stroke-width:1}
.ak-n.uretim rect.gov{stroke:var(--accent);stroke-width:1.6}
.ak-n.eksi  rect.gov{stroke:var(--critical);stroke-width:1.6}
.ak-n.sip   rect.gov{fill:var(--surface);stroke:var(--good);stroke-width:1.6}
.ak-n.sip.gec rect.gov{stroke:var(--critical)}
.ak-n text{fill:var(--ink);font-size:11.5px;font-weight:600}
.ak-n text.ad{fill:var(--ink2);font-size:9.5px;font-weight:400}
.ak-n text.et{fill:var(--muted);font-size:8px;font-weight:600;letter-spacing:.04em}
.ak-n text.sy{fill:var(--ink);font-size:10.5px;font-weight:600}
.ak-n text.sy.art{fill:var(--accent)}
.ak-n text.sy.eks{fill:var(--serious)}
.ak-n text.sy.kot{fill:var(--critical)}
.ak-k{fill:none;stroke:var(--axis);stroke-opacity:.75}
.ak-k.sip{stroke:var(--good)}
.ak-k.sip.gec{stroke:var(--critical)}
.ak-et{fill:var(--muted);font-size:9px}
.ak-bas{fill:var(--muted);font-size:10px;font-weight:600;letter-spacing:.06em}
.akis-sonuk .ak-n:not(.vurgu):not(.bulundu){opacity:.16}
.ak-n.bulundu rect.gov{stroke:var(--accent);stroke-width:2.5}
.ak-n.bulundu{filter:drop-shadow(0 0 7px color-mix(in srgb,var(--accent) 65%,transparent))}
.akis-sonuk .ak-k:not(.vurgu){opacity:.07}
.ak-n.vurgu rect.gov{stroke-width:2}
.legend{display:flex;align-items:center;gap:8px;flex-wrap:wrap;font-size:12px;color:var(--ink2);margin-top:12px}
.sw{display:flex;gap:2px}
.sw i{width:26px;height:11px;border-radius:2px;display:block}
.late{color:var(--critical);font-weight:600}
.ok{color:var(--good);font-weight:600}
.hide{display:none!important}
/* ayar paneli */
#ayar{position:fixed;inset:0;z-index:50;background:rgba(0,0,0,.45);
  display:flex;align-items:flex-start;justify-content:center;padding:28px 16px;overflow:auto}
#ayar[hidden]{display:none}
.panel{background:var(--surface);border:1px solid var(--ring);border-radius:12px;
  width:min(1080px,100%);box-shadow:0 18px 60px rgba(0,0,0,.3)}
.panel header{display:flex;align-items:center;gap:10px;padding:14px 18px;
  border-bottom:1px solid var(--grid);position:sticky;top:0;background:var(--surface);
  border-radius:12px 12px 0 0;z-index:2}
.panel header h2{font-size:16px;margin:0;flex:1}
.panel .body{padding:16px 18px 20px}
.panel footer{display:flex;gap:8px;align-items:center;padding:12px 18px;
  border-top:1px solid var(--grid);position:sticky;bottom:0;background:var(--surface);
  border-radius:0 0 12px 12px}
.tabs{display:flex;gap:4px;border-bottom:1px solid var(--grid);margin-bottom:14px;flex-wrap:wrap}
.tabs button{border:0;border-bottom:2px solid transparent;border-radius:0;background:none;
  padding:8px 12px;color:var(--ink2);font-weight:600}
.tabs button[aria-selected="true"]{color:var(--ink);border-bottom-color:var(--accent)}
.fld{display:grid;grid-template-columns:230px 1fr;gap:10px;align-items:center;
  padding:7px 0;border-bottom:1px solid var(--grid)}
.fld label{color:var(--ink2);font-size:13px}
.fld .hint2{font-size:11.5px;color:var(--muted);grid-column:2}
.fld input[type=number],.fld input[type=text],.fld input[type=date],.fld select{width:100%;max-width:330px}
.gun{display:flex;gap:3px;flex-wrap:wrap}
.gun label{display:flex;align-items:center;gap:3px;font-size:11.5px;border:1px solid var(--axis);
  border-radius:6px;padding:3px 7px;cursor:pointer;color:var(--ink)}
.gun input{width:auto;margin:0;padding:0}
.gun label:has(input:checked){background:var(--accent);border-color:var(--accent);color:#fff}
table.kay{border-collapse:collapse;width:100%;font-size:12.5px}
table.kay th{position:sticky;top:0;background:var(--surface);text-align:left;font-size:11px;
  color:var(--muted);text-transform:uppercase;letter-spacing:.04em;
  border-bottom:1px solid var(--axis);padding:6px}
table.kay td{border-bottom:1px solid var(--grid);padding:4px 6px}
table.kay input[type=number]{width:74px;padding:3px 5px}
table.kay tr.deg td{background:color-mix(in srgb,var(--accent) 12%,transparent)}
.kaylist{max-height:46vh;overflow:auto;border:1px solid var(--grid);border-radius:8px}
.durum{font-size:12px;color:var(--ink2)}
.durum.ok{color:var(--good);font-weight:600}
.durum.err{color:var(--critical);font-weight:600}
select.kisi{max-width:210px;padding:3px 6px;font-size:12px}
select.kisi.set{border-color:var(--accent);font-weight:600}
@media (max-width:640px){
  .wrap{padding:16px 16px 60px}
  table.hm th.row{min-width:150px;max-width:150px}
  .cell{height:46px}
}
</style>
</head>
<body>
<div class="wrap">
  <h1>Haftalık Üretim Çizelgesi</h1>
  <p class="sub" id="sub"></p>

  <div class="bar">
    <button id="prev" title="Önceki hafta">◀</button>
    <select id="hafta"></select>
    <button id="next" title="Sonraki hafta">▶</button>
    <select id="tip">
      <option value="">Tüm kaynak tipleri</option>
    </select>
    <input id="ara" type="search" placeholder="Kaynak veya ürün ara…" style="min-width:200px">
    <label class="hint"><input type="checkbox" id="bos" style="width:auto;padding:0;margin-right:4px"> boş kaynakları da göster</label>
    <span class="spacer"></span>
    <button id="tablo" aria-pressed="false">Tablo görünümü</button>
    <button id="ayarAc" title="Program ayarları">⚙ Ayarlar</button>
    <button id="tema" title="Açık / koyu tema">◐</button>
  </div>

  <div class="tiles" id="tiles"></div>

  <div class="tabs ana" id="anaTab" role="tablist">
    <button data-t="doluluk" aria-selected="true" role="tab">Kaynak doluluğu</button>
    <button data-t="cizelge" aria-selected="false" role="tab">Kaynak çizelgesi</button>
    <button data-t="emir" aria-selected="false" role="tab">İş emirleri</button>
    <button data-t="akis" aria-selected="false" role="tab">Sipariş akışı</button>
    <button data-t="gecikme" aria-selected="false" role="tab">Gecikmeler</button>
    <button data-t="kapatma" aria-selected="false" role="tab">Kaynak kapatma</button>
  </div>

  <section class="sek" data-t="doluluk">
  <div class="card">
    <h2>Kaynak × Gün — doluluk</h2>
    <p class="note">Her hücre o kaynağın o gün yüklenen saatini ve kapasite doluluğunu gösterir.
      Hücrenin üzerine gelince o gün hangi ürünlerin üretildiği çıkar; tıklayınca aşağıda
      satır satır listelenir. Taralı hücre = o gün çalışılmıyor.</p>
    <div id="hm"></div>
    <div class="legend">
      <span>Doluluk</span>
      <span class="sw" id="ramp"></span>
      <span>%0 → %100</span>
      <span style="margin-left:14px">·</span>
      <span>Hücredeki üst sayı = saat, alt sayı = doluluk %</span>
    </div>
  </div>

  <div class="card" id="grup"></div>
  <div class="det" id="detay"></div>
  </section>

  <section class="sek" data-t="cizelge" hidden>
    <div class="card" id="kcizelge"></div>
  </section>

  <section class="sek" data-t="emir" hidden>
    <div class="card" id="isler"></div>
  </section>

  <section class="sek" data-t="akis" hidden>
    <div class="card" id="akis"></div>
  </section>

  <section class="sek" data-t="gecikme" hidden>
    <div class="card" id="gecikme"></div>
  </section>

  <section class="sek" data-t="kapatma" hidden>
    <div class="card" id="kapatma"></div>
  </section>
</div>
<div id="tip" role="tooltip"></div>

<div id="ayar" hidden>
 <div class="panel" role="dialog" aria-modal="true" aria-label="Program ayarları">
  <header>
    <h2>Program ayarları</h2>
    <span class="durum" id="ayarDurum"></span>
    <button id="ayarKapat">Kapat</button>
  </header>
  <div class="body">
    <div class="tabs" id="ayarTabs">
      <button data-t="plan" aria-selected="true">Planlama</button>
      <button data-t="saat" aria-selected="false">Çalışma saatleri</button>
      <button data-t="op" aria-selected="false">Operatör</button>
      <button data-t="kay" aria-selected="false">Kaynak bazında</button>
    </div>
    <div id="t-plan"></div>
    <div id="t-saat" hidden></div>
    <div id="t-op" hidden></div>
    <div id="t-kay" hidden></div>
  </div>
  <footer>
    <button class="pri" id="ayarKaydet">ayarlar.json kaydet</button>
    <button id="ayarIndir">İndir</button>
    <button id="ayarSifirla">Varsayılanlara dön</button>
    <span class="spacer"></span>
    <span class="hint">Kaydettikten sonra <b>calistir.bat</b>&apos;ı çalıştır — plan yeni ayarlarla kurulur.</span>
  </footer>
 </div>
</div>

<script>
const V = /*__VERI__*/;
const GUNAD = ["Pzt","Sal","Çar","Per","Cum","Cmt","Paz"];
const TIPAD = {KISI:"Kişi", KALIP:"Kalıp", MAKINE:"Makine", DIGER:"Diğer"};
// 7 adimli tek-hue sekansiyel rampa: L monoton, komsu adim araligi >= 0.06
// (dataviz validator: lightness monotone PASS / adjacent dL PASS / single hue PASS).
// En acik adim "sifira yakin" demek; her hucre sayiyi da yazdigi icin renk tek
// basina anlam tasimiyor + tablo gorunumu var.
const RAMP = ["--s100","--s200","--s300","--s400","--s500","--s600","--s700"];
const INK  = ["--oi1","--oi2","--oi3","--oi4","--oi5","--oi6","--oi7"];
const rampIx = p => Math.min(RAMP.length-1, Math.max(0, Math.round(p/100*(RAMP.length-1))));

const MAL = V.malzemeler, KYN = V.kaynaklar;
const MALIX = new Map(MAL.map((m,i)=>[m[0], i]));
const malAd = kod => { const i = MALIX.get(kod); return i==null ? "" : (MAL[i][1]||""); };
let gecHepsi = false;
// ---- kaynak kapatmalari (izin / ariza / bakim)
let KAPANIS = (V.kapanis || []).map(k => ({
  kaynak: String(k.kaynak||""), bas: k.bas||"", bit: k.bit||k.bas||"",
  tur: k.tur||"IZIN", oran: Number(k.oran||0), not: k.not||""}));
// o kaynagin o gunku kalan kapasite orani (1 = acik, 0 = kapali)
function kapOran(kod, g){
  let o = 1;
  for(const k of KAPANIS)
    if(k.kaynak===kod && k.bas<=g && g<=k.bit) o = Math.min(o, (k.oran||0)/100);
  return o;
}
function kapNeden(kod, g){
  for(const k of KAPANIS)
    if(k.kaynak===kod && k.bas<=g && g<=k.bit)
      return k.tur + (k.not ? " — "+k.not : "") + (k.oran>0 ? ` (%${k.oran} kapasite)` : "");
  return "";
}
const fmt = (n,d=0)=>n.toLocaleString("tr-TR",{minimumFractionDigits:d,maximumFractionDigits:d});
const d2 = s => { const [y,m,g]=s.split("-").map(Number); return new Date(y,m-1,g); };
// DIKKAT: toISOString() UTC'ye cevirir; UTC+ saat dilimlerinde yerel geceyarisi
// bir onceki gune duser ve tum hafta anahtarlari kayar. Yerel tarih parcalarini
// kullanmak sart.
const iso = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
function pzt(d){ const x=new Date(d); const w=(x.getDay()+6)%7; x.setDate(x.getDate()-w); return x; }
const trt = d => d.toLocaleDateString("tr-TR",{day:"2-digit",month:"short"});

// ---- kayitlari haftalara dagit
// Bir uretim kaydi iki satira birden girer: calistigi KAYNAK (kisi/kalip/firin)
// ve -varsa- uzerinde calistigi MAKINE. Saat ikisinde de ayni, cunku kalip ve
// makine es zamanli tukenir. Toplam saat sayilirken makine satiri sayilmaz.
const H = new Map();                     // haftaBasi -> gun -> kaynakIx -> [kayit]
function koy(hb,g,ri,k){
  if(!H.has(hb)) H.set(hb,new Map());
  const w=H.get(hb); if(!w.has(g)) w.set(g,new Map());
  const d=w.get(g); if(!d.has(ri)) d.set(ri,[]); d.get(ri).push(k);
}
for (const k of V.kayitlar){
  const hb = iso(pzt(d2(k[0])));
  koy(hb, k[0], k[2], k);
  if(k[3] >= 0) koy(hb, k[0], k[3], k);
}
// motorun o kaynak-gun icin kaydettigi gercek yuk (havuzlar icin tek dogru kaynak)
const motorYuk = (ki,g) => (V.yuk[KYN[ki][0]]||{})[g];
// kaynagin o gunku kapasitesi: takvimden (saat/gun x o gun calisiliyor mu)
function kapasite(ki, g){
  const K = KYN[ki]; if(!K) return 0;
  const t = K[4][(d2(g).getDay()+6)%7] ? K[3] : 0;
  return t * kapOran(K[0], g);
}
const HAFTALAR = [...new Set([...H.keys(),
  ...Object.keys((V.akis&&V.akis.hafta)||{})])].sort();
let hi = 0, secili = null, tabloMod = false;

// ---- kontroller
const el = id => document.getElementById(id);
const selH = el("hafta"), selT = el("tip");
HAFTALAR.forEach((h,i)=>{
  const b = pzt(d2(h)), s = new Date(b); s.setDate(s.getDate()+6);
  const o = document.createElement("option");
  o.value = i; o.textContent = `${i+1}. hafta — ${trt(b)} – ${trt(s)} ${s.getFullYear()}`;
  selH.appendChild(o);
});
[...new Set(KYN.map(k=>k[2]))].sort().forEach(t=>{
  const o=document.createElement("option"); o.value=t; o.textContent=TIPAD[t]||t; selT.appendChild(o);
});
el("ramp").innerHTML = RAMP.map(r=>`<i style="background:var(${r})"></i>`).join("");

el("prev").onclick = ()=>{ hi=Math.max(0,hi-1); selH.value=hi; ciz(); };
el("next").onclick = ()=>{ hi=Math.min(HAFTALAR.length-1,hi+1); selH.value=hi; ciz(); };
selH.onchange = ()=>{ hi=+selH.value; ciz(); };
selT.onchange = el("ara").oninput = el("bos").onchange = ()=>ciz();
el("tablo").onclick = e=>{ tabloMod=!tabloMod; e.target.setAttribute("aria-pressed",tabloMod); ciz(); };
el("tema").onclick = ()=>{
  const d = document.documentElement;
  const koyu = d.getAttribute("data-theme")==="dark" ||
    (!d.getAttribute("data-theme") && matchMedia("(prefers-color-scheme: dark)").matches);
  d.setAttribute("data-theme", koyu ? "light" : "dark");
};

// ---- tooltip
const tip = el("tip");
function tipGoster(e, html){
  tip.innerHTML = html; tip.classList.add("on");
  const r = tip.getBoundingClientRect();
  let x = e.clientX+14, y = e.clientY+14;
  if(x+r.width > innerWidth-8) x = e.clientX-r.width-14;
  if(y+r.height > innerHeight-8) y = Math.max(8, e.clientY-r.height-14);
  tip.style.left = x+"px"; tip.style.top = y+"px";
}
const tipKapat = ()=>tip.classList.remove("on");

const rampRenk = p => p<=0 ? "var(--grid)" : `var(${RAMP[rampIx(p)]})`;
const rampInk  = p => p<=0 ? "var(--ink)"   : `var(${INK[rampIx(p)]})`;

// ---- ana cizim
function ciz(){
  const hb = HAFTALAR[hi];
  const gunler = [...Array(7)].map((_,i)=>{ const d=pzt(d2(hb)); d.setDate(d.getDate()+i); return iso(d); });
  const w = H.get(hb) || new Map();
  const tipF = selT.value, q = el("ara").value.trim().toLocaleLowerCase("tr");

  // kaynak -> gun -> {saat, adet, kayitlar}
  const sat = new Map();
  for(const g of gunler){
    const gm = w.get(g); if(!gm) continue;
    for(const [ki,lst] of gm){
      const uy = lst.filter(k => !q ||
        (KYN[ki][0]+" "+KYN[ki][1]).toLocaleLowerCase("tr").includes(q) ||
        (MAL[k[1]][0]+" "+MAL[k[1]][1]).toLocaleLowerCase("tr").includes(q));
      if(!uy.length) continue;
      if(!sat.has(ki)) sat.set(ki, new Map());
      sat.get(ki).set(g, {
        saat: uy.reduce((a,k)=>a+k[5],0),
        adet: uy.reduce((a,k)=>a+k[4],0),
        kayitlar: uy});
    }
  }
  let kaynakIdx = [...Array(KYN.length).keys()]
    .filter(i => !tipF || KYN[i][2]===tipF)
    .filter(i => el("bos").checked || sat.has(i) || Object.keys(V.yuk[KYN[i][0]]||{})
                   .some(g => gunler.includes(g)));
  if(q) kaynakIdx = kaynakIdx.filter(i => sat.has(i));

  // --- stat tiles
  let topSaat=0, topAdet=0, enDolu=null;
  const urunSet = new Set();
  for(const [ki,gm] of sat) for(const [g,v] of gm){
    if(KYN[ki][2]!=="MAKINE" || v.kayitlar.every(k=>k[3]<0)){ topSaat+=v.saat; topAdet+=v.adet; }
    v.kayitlar.forEach(k=>urunSet.add(k[1]));
    const kp = kapasite(ki,g);
    const p = kp? Math.min(100, v.saat/kp*100) : 0;
    if(!enDolu || p>enDolu.p) enDolu={ki,g,p,saat:v.saat,kap:kp};
  }
  const bas = pzt(d2(hb)), son = new Date(bas); son.setDate(son.getDate()+6);
  el("sub").textContent =
    `${HAFTALAR.length} haftalık plan · ${fmt(V.kayitlar.length)} üretim kaydı · ` +
    `plan ${V.ozet["Plan baslangic tarihi"]||""} · oluşturuldu ${V.uretildi}`;
  el("tiles").innerHTML = [
    ["Hafta", `${trt(bas)} – ${trt(son)}`, son.getFullYear()],
    ["Yüklenen saat", fmt(topSaat,1), `${sat.size} kaynak çalışıyor`],
    ["Üretilen adet", fmt(topAdet), `${urunSet.size} farklı ürün`],
    ["En dolu kaynak", enDolu? `%${fmt(enDolu.p,1)}` : "–",
      enDolu? `${KYN[enDolu.ki][0]} — ${KYN[enDolu.ki][1]||""}`.slice(0,40) : ""],
  ].map(([l,v,f])=>`<div class="tile"><div class="lbl">${l}</div><div class="val">${v}</div><div class="foot">${f||""}</div></div>`).join("");

  // --- tablo gorunumu
  if(tabloMod){
    const rows=[];
    for(const g of gunler){
      const gm = w.get(g); if(!gm) continue;
      for(const [ki,lst] of gm){
        if(tipF && KYN[ki][2]!==tipF) continue;
        if(KYN[ki][2]==="MAKINE" && lst.some(k=>k[3]>=0 && k[3]!==ki)) continue; // ayni kayit iki kez cikmasin
        for(const k of lst){
          if(q && !((KYN[ki][0]+" "+KYN[ki][1]+" "+MAL[k[1]][0]+" "+MAL[k[1]][1])
              .toLocaleLowerCase("tr").includes(q))) continue;
          rows.push(`<tr><td>${g}</td><td>${GUNAD[(d2(g).getDay()+6)%7]}</td>
            <td><b>${KYN[ki][0]}</b><br><span class="rn">${KYN[ki][1]||""}</span></td>
            <td><span class="pill">${TIPAD[KYN[ki][2]]||KYN[ki][2]}</span></td>
            <td>${k[3]>=0?KYN[k[3]][0]:"–"}</td>
            <td>${MAL[k[1]][0]}<br><span class="rn">${MAL[k[1]][1]}</span></td>
            <td class="n">${fmt(k[4])}</td><td class="n">${fmt(k[5],2)}</td></tr>`);
        }
      }
    }
    el("hm").innerHTML = `<table class="dt"><thead><tr>
      <th>Tarih</th><th>Gün</th><th>Kaynak</th><th>Tip</th><th>Makine</th>
      <th>Ürün</th><th>Adet</th><th>Saat</th></tr></thead>
      <tbody>${rows.join("") || `<tr><td colspan="8">Bu haftada kayıt yok.</td></tr>`}</tbody></table>`;
  } else {
    // --- isi haritasi
    const grup = {};
    kaynakIdx.forEach(i=>{ (grup[KYN[i][2]] ||= []).push(i); });
    const sira = ["KISI","KALIP","MAKINE","DIGER"].filter(t=>grup[t]);
    Object.keys(grup).forEach(t=>{ if(!sira.includes(t)) sira.push(t); });
    let h = `<table class="hm"><thead><tr><th class="row">Kaynak</th>`;
    gunler.forEach(g=>{ const d=d2(g);
      h += `<th>${GUNAD[(d.getDay()+6)%7]}<br><span class="rn">${trt(d)}</span></th>`; });
    h += `<th>Hafta</th></tr></thead><tbody>`;
    for(const t of sira){
      const list = grup[t].sort((a,b)=>{
        const sa = [...(sat.get(a)||new Map()).values()].reduce((x,v)=>x+v.saat,0);
        const sb = [...(sat.get(b)||new Map()).values()].reduce((x,v)=>x+v.saat,0);
        return sb-sa || KYN[a][0].localeCompare(KYN[b][0],"tr");
      });
      h += `<tr class="grp"><th class="row">${TIPAD[t]||t} · ${list.length}</th>`+
           `<td colspan="8"></td></tr>`;
      for(const ki of list){
        h += `<tr><th class="row"><span class="rk">${KYN[ki][0]}</span>`+
             `<span class="rn">${KYN[ki][1]||""}</span></th>`;
        let hs=0;
        for(const g of gunler){
          let v = (sat.get(ki)||new Map()).get(g);
          const kp = kapasite(ki,g);
          const my = motorYuk(ki,g);
          // havuz kaynaklar (orn. operator) UretimPlani'nda kaynak satiri olarak
          // gecmez; yuklerini motorun kendi kaydindan al
          if(!v && my) v = {saat: my, adet: 0, kayitlar: []};
          if(!v){
            const nd = kapNeden(KYN[ki][0], g);
            h += `<td><div class="cell ${kp>0?"empty":(nd?"off kapali":"off")}"
              title="${kp>0?"boş":(nd||"çalışılmıyor")}">${nd&&kp<=0?"✕":""}</div></td>`;
            continue;
          }
          hs += v.saat;
          const p = kp? Math.min(100, v.saat/kp*100) : 0;
          h += `<td><button class="cell${secili&&secili.ki===ki&&secili.g===g?" sel":""}"
            style="background:${rampRenk(p)};color:${rampInk(p)}"
            data-ki="${ki}" data-g="${g}">
            <span class="h">${fmt(v.saat,1)}</span>
            <span class="p">%${fmt(p,0)}</span></button></td>`;
        }
        h += `<td class="tot"><div class="cell"><span class="h">${fmt(hs,1)}</span>
              <span class="p">saat</span></div></td></tr>`;
      }
    }
    h += `</tbody></table>`;
    if(!kaynakIdx.length) h = `<p class="note">Bu filtreyle gösterilecek kaynak yok.</p>`;
    el("hm").innerHTML = h;

    el("hm").querySelectorAll("button.cell").forEach(b=>{
      const ki=+b.dataset.ki, g=b.dataset.g;
      const v = (sat.get(ki)||new Map()).get(g)
        || {saat: motorYuk(ki,g)||0, adet:0, kayitlar:[]};
      const kp = kapasite(ki,g);
      const p = kp? Math.min(100, v.saat/kp*100) : 0;
      const top = [...v.kayitlar].sort((a,b)=>b[5]-a[5]);
      b.onmousemove = e => tipGoster(e,
        `<b>${KYN[ki][0]}</b> ${KYN[ki][1]?`<span class="m">— ${KYN[ki][1]}</span>`:""}
         <div class="r"><span class="m">${g} ${GUNAD[(d2(g).getDay()+6)%7]}</span>
           <span>${fmt(v.saat,2)} / ${fmt(kp,1)} sa · %${fmt(p,0)}</span></div><hr>`+
        top.slice(0,7).map(k=>`<div class="r"><span><b>${MAL[k[1]][0]}</b>${k[3]>=0?` <span class="m">(${KYN[ki][2]==="MAKINE"?KYN[k[2]][0]:KYN[k[3]][0]})</span>`:""}<br><span class="m">${kis(MAL[k[1]][1]||"",44)}</span></span>
           <span>${fmt(k[4])} ad · ${fmt(k[5],2)} sa</span></div>`).join("")+
        (top.length>7?`<div class="r m">+${top.length-7} ürün daha…</div>`:"")+
        `<hr><div class="r m">Satır satır görmek için tıkla</div>`);
      b.onmouseleave = tipKapat;
      b.onclick = ()=>{ secili={ki,g}; ciz();
        el("detay").scrollIntoView({behavior:"smooth",block:"nearest"}); };
    });
  }

  // --- detay
  if(secili && (sat.get(secili.ki)?.get(secili.g) || motorYuk(secili.ki,secili.g))){
    const {ki,g} = secili;
    const v = sat.get(ki)?.get(g) || {saat: motorYuk(ki,g)||0, adet:0, kayitlar:[]};
    const kp = kapasite(ki,g);
    el("detay").classList.remove("hide");
    el("detay").innerHTML = `<h2 style="font-size:14px;margin:0 0 3px">
        ${KYN[ki][0]} — ${g} ${GUNAD[(d2(g).getDay()+6)%7]}</h2>
      <p class="note">${KYN[ki][1]||""} · <span class="pill">${TIPAD[KYN[ki][2]]||KYN[ki][2]}</span>
        · ${fmt(v.saat,2)} / ${fmt(kp,1)} saat (%${fmt(kp?Math.min(100,v.saat/kp*100):0,0)})
        · ${fmt(v.adet)} adet · ${v.kayitlar.length} üretim satırı</p>
      <table class="dt"><thead><tr><th>Ürün</th><th>Açıklama</th><th>Makine</th>
        <th class="n">Adet</th><th class="n">Saat</th></tr></thead><tbody>`+
      [...v.kayitlar].sort((a,b)=>b[5]-a[5]).map(k=>
        `<tr><td><b>${MAL[k[1]][0]}</b></td><td>${MAL[k[1]][1]}</td>
          <td>${k[3]>=0 ? (KYN[ki][2]==="MAKINE"?KYN[k[2]][0]:KYN[k[3]][0]) : "–"}</td><td class="n">${fmt(k[4])}</td>
          <td class="n">${fmt(k[5],2)}</td></tr>`).join("")+
      `</tbody></table>`;
  } else {
    el("detay").classList.add("hide");
  }

  // --- GRUP (havuz) kaynaklarin dolulugu: paralel adet > 1 olanlar
  // (a) paralel adedi >1 olan havuz kaynaklar (orn. ENJ-OPERATOR)
  // (b) uyeleri ayri kaynak olarak atanmis kapasite gruplari (MONTAJ, BILEZIK...)
  const havuzlar = Object.keys(V.kaynakAyar||{}).filter(k=>
        (V.kaynakAyar[k].paralel||1) > 1 && (V.gruplar[k]||[]).length > 0);
  // Havuz olarak modellenen gruplar (orn. ENJEKSIYON = ENJ-OPERATOR havuzu) bu
  // tabloda tekrar gosterilmez; yukleri havuz kartinda zaten var. Bir kisi birden
  // fazla kapasite grubunda olabildigi icin havuz uye listesi gruptan birkac kisi
  // eksik olabilir (orn. bireysel planlanan operator) - bu yuzden kume esitligi
  // degil "havuz uyeleri grubun icinde mi" testi kullaniliyor.
  const havuzKapsar = g => {
    const uye = new Set(V.gruplar[g]||[]);
    return havuzlar.some(h => {
      const hu = V.gruplar[h]||[];
      return hu.length > 0 && hu.every(k => uye.has(k));
    });
  };
  const kisiGruplari = Object.keys(V.gruplar||{}).filter(g=>
        !havuzlar.includes(g) && !havuzKapsar(g) &&
        (V.gruplar[g]||[]).length > 1 &&
        V.gruplar[g].some(k => V.kaynakAyar[k]));
  let gh = "";
  for(const hk of havuzlar){
    const ki = KYN.findIndex(x=>x[0]===hk);
    if(ki < 0) continue;
    const d = V.kaynakAyar[hk];
    const uyeler = (V.gruplar[hk]||[]);
    let topY=0, topK=0, sat3=[];
    for(const g of gunler){
      const kp = kapasite(ki,g), y = motorYuk(ki,g) || 0;
      topY+=y; topK+=kp;
      sat3.push({g,kp,y, esd: d.saat? y/d.saat : 0});
    }
    gh += `<h2>${hk} — havuz doluluğu</h2>
      <p class="note">${d.ad||""} · ${d.paralel} kişi × ${fmt(d.saat,1)} sa/gün
        · hafta toplamı ${fmt(topY,1)} / ${fmt(topK,1)} saat
        (%${fmt(topK? topY/topK*100:0,1)})</p>
      <table class="dt"><thead><tr><th>Gün</th><th class="n">Yük (saat)</th>
        <th class="n">Kapasite (saat)</th><th class="n">Doluluk %</th>
        <th class="n">Gereken kişi</th><th class="n">Boşta kalan</th></tr></thead><tbody>`+
      sat3.map(x=>`<tr><td>${x.g} ${GUNAD[(d2(x.g).getDay()+6)%7]}</td>
        <td class="n">${fmt(x.y,2)}</td><td class="n">${fmt(x.kp,1)}</td>
        <td class="n">${x.kp? fmt(Math.min(100,x.y/x.kp*100),0):"–"}</td>
        <td class="n"><b>${x.kp? Math.ceil(x.esd-1e-9):0}</b> / ${d.paralel}</td>
        <td class="n">${x.kp? Math.max(0, d.paralel-Math.ceil(x.esd-1e-9)) : "–"}</td></tr>`).join("")+
      `</tbody></table>
      <p class="hint" style="margin-top:8px">Gereken kişi = o günkü yük ÷ ${fmt(d.saat,1)} saat,
        yukarı yuvarlanmış. Havuzdaki ${uyeler.length} kişi:
        ${uyeler.slice(0,10).map(k=>{const i=KYN.findIndex(x=>x[0]===k);
          return `<span class="pill">${k}${i>=0&&KYN[i][1]?" "+KYN[i][1].split(" - ")[0]:""}</span>`;}).join(" ")}
        ${uyeler.length>10?`<span class="pill">+${uyeler.length-10}</span>`:""}</p>`;
  }
  // --- uyeleri tek tek atanan kapasite gruplari
  // Grubun KENDI isi: kapasite grubu o grup olan uretim kayitlarinin saati.
  // (Uye yuklerini toplamak yanlis olurdu: paketleme gruplari ayni kisileri paylasiyor.)
  const grupYuk = new Map();            // grup -> gun -> saat
  for(const g of gunler){
    const gm = w.get(g); if(!gm) continue;
    for(const [ki,lst] of gm){
      if(KYN[ki][2]==="MAKINE" && lst.some(k=>k[3]>=0 && k[3]!==ki)) continue;
      for(const k of lst){
        const kg = k[6]; if(!kg) continue;
        if(!grupYuk.has(kg)) grupYuk.set(kg, new Map());
        const m2 = grupYuk.get(kg);
        m2.set(g, (m2.get(g)||0) + k[5]);
      }
    }
  }
  const gsat = [];
  for(const g of kisiGruplari){
    const uy = V.gruplar[g].filter(k=>V.kaynakAyar[k]);
    let topY=0, topK=0, uyeY=0, gunluk=[];
    for(const gn of gunler){
      let kp=0, mesgul=0, uyy=0;
      for(const k of uy){
        const ki = KYN.findIndex(x=>x[0]===k);
        const d = V.kaynakAyar[k];
        kp += d.gunler[(d2(gn).getDay()+6)%7] ? d.saat*(d.paralel||1) : 0;
        const yy = ki>=0 ? (motorYuk(ki,gn)||0) : 0;
        uyy += yy; if(yy>0.001) mesgul++;
      }
      const y = (grupYuk.get(g)||new Map()).get(gn) || 0;
      topY+=y; topK+=kp; uyeY+=uyy;
      gunluk.push({gn,y,kp,mesgul});
    }
    if(topY<=0.001 && !el("bos").checked) continue;
    gsat.push({g, uy, topY, topK, uyeY, gunluk, pkt: g.toUpperCase().startsWith("PKT-GRUP")});
  }
  gsat.sort((a,b)=> (a.pkt-b.pkt) || ((b.topK? b.topY/b.topK:0)-(a.topK? a.topY/a.topK:0)));
  if(gsat.length){
    gh += `<h2 style="margin-top:${gh?"22px":"0"}">Kapasite gruplarının doluluğu</h2>
      <p class="note">Üyeleri tek tek atanan gruplar — ara işlemler (montaj, bilezik, çapak,
        lebleme, pim çakma…) ve paketleme. Gün hücrelerindeki saat o grubun kendi işi;
        altındaki sayı o gün en az bir iş alan üye sayısı.</p>
      <table class="dt"><thead><tr><th>Kapasite grubu</th><th class="n">Kişi</th>`+
      gunler.map(gn=>`<th class="n">${GUNAD[(d2(gn).getDay()+6)%7]}<br><span class="rn">${trt(d2(gn))}</span></th>`).join("")+
      `<th class="n">Grubun işi</th><th class="n">Üyelerin doluluğu</th></tr></thead><tbody>`+
      gsat.map(x=>`<tr><td><b>${x.g}</b>${x.pkt?' <span class="pill">paketleme</span>':""}
          <br><span class="rn">${x.uy.slice(0,4).join(", ")}${x.uy.length>4?" +"+(x.uy.length-4):""}</span></td>
        <td class="n">${x.uy.length}</td>`+
        x.gunluk.map(d=>`<td class="n">${d.kp? `${fmt(d.y,1)}<br><span class="rn">${d.mesgul}/${x.uy.length} kişi</span>`:"–"}</td>`).join("")+
        `<td class="n"><b>${fmt(x.topY,1)}</b> sa<br><span class="rn">%${fmt(x.topK? x.topY/x.topK*100:0,0)} kapasite</span></td>
         <td class="n"><b>%${fmt(x.topK? x.uyeY/x.topK*100:0,0)}</b><br><span class="rn">${fmt(x.uyeY,1)} / ${fmt(x.topK,0)} sa</span></td></tr>`).join("")+
      `</tbody></table>
      <p class="hint" style="margin-top:8px"><b>Grubun işi</b> = kapasite grubu bu olan üretimlerin saati.
        <b>Üyelerin doluluğu</b> = aynı kişilerin tüm işlerdeki toplam yükü — paketleme grupları
        aynı kişileri paylaştığı için bu oran grubun kendi işinden yüksek olabilir.</p>`;
  }
  el("grup").innerHTML = gh || `<h2>Grup kaynak yok</h2>
    <p class="note">Bu haftada grup olarak atanmış çalışan kaynak bulunmuyor.</p>`;

  // --- yapilacak isler (toplulastirilmis is emirleri)
  const em = (V.emirler||[]).filter(e =>
        e[4] <= gunler[6] && e[5] >= gunler[0] &&
        (!tipF || KYN[e[2]][2]===tipF || (e[3]>=0 && KYN[e[3]][2]===tipF)) &&
        (!q || (MAL[e[1]][0]+" "+MAL[e[1]][1]+" "+KYN[e[2]][0]+" "+KYN[e[2]][1]+
                (e[3]>=0?" "+KYN[e[3]][0]:"")).toLocaleLowerCase("tr").includes(q)))
      .sort((a,b)=> a[4].localeCompare(b[4]) || b[8]-a[8]);
  const emSaat = em.reduce((a,e)=>a+e[8],0), emAdet = em.reduce((a,e)=>a+e[7],0);
  el("isler").innerHTML = `<h2>Yapılacak işler — iş emirleri</h2>
    <p class="note">${fmt(em.length)} iş emri · ${fmt(emAdet)} adet · ${fmt(emSaat,1)} saat.
      Aynı ürün hafta boyunca aynı kaynakta devam ediyorsa tek iş emrinde toplandı;
      kaynak değişirse ya da hafta dönerse yeni emir açılır.</p>`+
    (em.length ? `<table class="dt"><thead><tr>
        <th>İş Emri</th><th>Ürün</th><th>Kaynak</th><th>Makine</th>
        <th>Kapasite grubu</th><th>Kişi grubu</th><th>Gerçekte kim yapacak</th>
        <th>Başlangıç</th><th>Bitiş</th><th class="n">Gün</th>
        <th class="n">Miktar</th><th class="n">Saat</th><th>Tetikleyen sipariş</th>
      </tr></thead><tbody>`+
      em.map(e=>{
        const cokGun = e[6]>1;
        return `<tr><td class="nw"><b>${e[0]}</b></td>
          <td>${MAL[e[1]][0]}<br><span class="rn">${MAL[e[1]][1]}</span></td>
          <td class="nw">${KYN[e[2]][0]} <span class="pill">${TIPAD[KYN[e[2]][2]]||KYN[e[2]][2]}</span>
              ${e[12]>1?`<span class="pill warn">${e[12]} kişilik ekip</span>`:""}
              <br><span class="rn">${KYN[e[2]][1]||""}</span>
              ${(e[13]||[]).length?`<br><span class="rn">+ ${e[13].map(i=>KYN[i][0]+(KYN[i][1]?" "+KYN[i][1].split(" - ")[0]:"")).join(" · ")}</span>`:""}</td>
          <td class="nw">${e[3]>=0?KYN[e[3]][0]:"–"}</td>
          <td class="nw">${e[11]? `<span class="pill">${e[11]}</span>`:"–"}</td>
          <td class="nw">${e[10]? `<span class="pill">${e[10]}</span>`:"–"}</td>
          <td>${kisiSecici(e)}</td>
          <td class="nw">${e[4]}</td><td class="nw">${e[5]}</td>
          <td class="n">${e[6]}${cokGun?' <span class="pill">toplu</span>':""}</td>
          <td class="n">${fmt(e[7])}</td><td class="n">${fmt(e[8],2)}</td>
          <td><span class="rn">${e[9]}</span></td></tr>`;}).join("")+
      `</tbody></table>
       <div class="bar" style="position:static;margin-top:12px">
         <button class="pri" id="ataKaydet">atamalar.json kaydet</button>
         <button id="ataTemizle">Bu haftanın atamalarını temizle</button>
         <span class="durum" id="ataDurum">${Object.keys(ATAMA).length} atama kayıtlı</span>
         <span class="spacer"></span>
         <span class="hint">Seçim planı değiştirmez; kimin yapacağını kaydeder.</span>
       </div>`
    : `<p class="note">Bu haftada iş emri yok.</p>`);
  if(em.length){
    ataDinle();
    el("ataKaydet").onclick = ()=> dosyaKaydet("atamalar.json", ATAMA, el("ataDurum"));
    el("ataTemizle").onclick = ()=>{ em.forEach(e=>delete ATAMA[e[0]]); ciz(); };
  }

  // --- o hafta teslim tarihi gecen siparisler
  const gecH = V.geciken.filter(x => x.teslim >= gunler[0] && x.teslim <= gunler[6]);
  const gecL = gecHepsi ? V.geciken.slice() : gecH;
  const gecTablo = lst => `<table class="dt"><thead><tr><th>Belge</th><th>Kalem</th>
      <th>Ürün</th><th>Açıklama</th>
      <th class="n">Miktar</th><th>Teslim</th><th>Planlanan bitiş</th>
      <th class="n">Gecikme</th></tr></thead><tbody>`+
    lst.slice().sort((a,b)=>b.gun-a.gun).map(x=>`<tr><td>${x.belge}</td><td>${x.kalem}</td>
      <td><b>${x.malzeme}</b></td><td>${malAd(x.malzeme)}</td>
      <td class="n">${fmt(x.miktar)}</td><td>${x.teslim}</td>
      <td>${x.bitis}</td><td class="n late">+${x.gun} gün</td></tr>`).join("")+
    `</tbody></table>`;
  el("gecikme").innerHTML =
    `<h2>${gecHepsi ? "Planın tamamındaki geciken siparişler"
                    : "Bu hafta teslim edilmesi gereken ama geciken siparişler"}</h2>
     <p class="note">${gecL.length} kalem · bu hafta ${gecH.length},
       planın tamamında ${V.geciken.length}.</p>
     <div class="akis-bar">
       <label><input type="checkbox" id="gecHep" ${gecHepsi?"checked":""}>
         planın tamamını göster</label>
     </div>`
    + (gecL.length ? gecTablo(gecL)
       : `<p class="note ok">Bu haftaya teslim tarihi düşen siparişlerin tamamı zamanında.</p>`);
  if(el("gecHep")) el("gecHep").onchange = ev => { gecHepsi = ev.target.checked; ciz(); };

  cizAkis(hb);
  cizKaynakCizelge(hb, gunler, em);
  cizKapatma();

  // sekme rozetleri
  const A0 = (AKIS.hafta||{})[hb];
  const roz = {emir: em.length, gecikme: gecH.length, kapatma: KAPANIS.length,
               akis: A0 ? A0.n.length : 0,
               cizelge: new Set(em.flatMap(e =>
                 [e[2]].concat(e[3]>=0?[e[3]]:[]).concat(e[13]||[]))).size};
  document.querySelectorAll("#anaTab button").forEach(b=>{
    const v = roz[b.dataset.t];
    const ad = b.dataset.ad || (b.dataset.ad = b.textContent.trim());
    b.innerHTML = ad + (v != null ? ` <span class="sy">${v}</span>` : "");
  });
}


// ===================== AYARLAR =====================
const GUNAD7 = ["Pzt","Sal","Çar","Per","Cum","Cmt","Paz"];
const P = V.param || {}, KA = V.kaynakAyar || {};
function sayi(x, y){ const n = parseFloat(String(x ?? "").replace(",", ".")); return isFinite(n)? n : y; }

// kayitli ayarlar varsa onu, yoksa plan girdisindeki mevcut degerleri baz al
let A = JSON.parse(JSON.stringify(V.kayitliAyar && Object.keys(V.kayitliAyar).length
  ? V.kayitliAyar : {}));
A.plan = Object.assign({
  baslangic: P["Plan Baslangic Tarihi"] || "",
  aktarma: P["Aktarma Kurali"] || "GUNLUK",
  aktarma_partisi: sayi(P["Aktarma Partisi"], 0),
  ufuk: sayi(P["Planlama Ufku (gun)"], 1100),
  siralama: P["Siralama Kurali"] || "EDD",
  gunluk_is_cesidi: sayi(P["Gunluk Is Cesidi"], 1),
  is_basina_kaynak: sayi(P["Is Basina Azami Kaynak"], 1),
  takviye_payi: sayi(P["Takviye Emniyet Payi (gun)"], 2),
  erken_sevkiyat: P["Erken Sevkiyat"] || "HAYIR",
  min_kalip_kalis: sayi(P["Minimum Kalip Kalis (gun)"], 0),
  tatiller: P["Tatiller"] || "",
  senaryo: P["Senaryo Analizi"] || "HAYIR",
  satinalma: P["Satinalma Kisiti"] || "HAYIR",
}, A.plan || {});
// varsayilan calisma saatlerini mevcut kaynaklardan tahmin et
function ornek(filtre, vs){
  for(const [k,d] of Object.entries(KA)) if(filtre(k,d)) return {saat:d.saat, gunler:d.gunler};
  return vs;
}
A.varsayilan = Object.assign({
  kisi:     ornek((k,d)=>d.tip==="KISI" && k!=="ENJ-OPERATOR", {saat:9,gunler:[1,1,1,1,1,0,0]}),
  operator: ornek((k)=>k==="ENJ-OPERATOR", {saat:8,gunler:[1,1,1,1,1,1,0]}),
  makine:   ornek((k,d)=>d.tip==="MAKINE" && !k.startsWith("FIRIN"), {saat:24,gunler:[1,1,1,1,1,1,0]}),
  kalip:    ornek((k,d)=>d.tip==="KALIP", {saat:24,gunler:[1,1,1,1,1,1,0]}),
  firin:    ornek((k)=>k.startsWith("FIRIN"), {saat:24,gunler:[1,1,1,1,1,1,0]}),
}, A.varsayilan || {});
A.operator = Object.assign({
  katsayi: sayi((KA["KM08"]||{}).ekKats, 1) || 1,
  kisit: Object.values(KA).some(d=>d.ek==="ENJ-OPERATOR"),
}, A.operator || {});
A.kaynaklar = A.kaynaklar || {};
const A0 = JSON.parse(JSON.stringify(A));

const fld = (l,i,h)=>`<div class="fld"><label>${l}</label><div>${i}</div>${h?`<div class="hint2">${h}</div>`:""}</div>`;
const gunKutu = (ad,g)=>`<div class="gun">`+GUNAD7.map((n,i)=>
  `<label><input type="checkbox" data-gun="${ad}" data-i="${i}" ${g[i]?"checked":""}>${n}</label>`).join("")+`</div>`;

function ayarCiz(){
  el("t-plan").innerHTML =
    fld("Plan başlangıç tarihi", `<input type="date" data-p="baslangic" value="${A.plan.baslangic||""}">`,
        "İleri planlamanın başladığı gün") +
    fld("Aktarma kuralı", `<select data-p="aktarma">`+
        ["GUNLUK","ADET","SABIT"].map(x=>`<option ${A.plan.aktarma===x?"selected":""}>${x}</option>`).join("")+
        `</select>`, "GUNLUK = alt seviyenin 1 günlük üretimi birikince üst seviye başlar") +
    fld("Aktarma partisi (adet)", `<input type="number" min="0" data-p="aktarma_partisi" value="${A.plan.aktarma_partisi}">`,
        "Yalnızca aktarma kuralı SABIT ise geçerli") +
    fld("Planlama ufku (gün)", `<input type="number" min="30" data-p="ufuk" value="${A.plan.ufuk}">`) +
    fld("Sıralama kuralı", `<select data-p="siralama">`+
        ["EDD","FIFO"].map(x=>`<option ${A.plan.siralama===x?"selected":""}>${x}</option>`).join("")+
        `</select>`, "EDD = en erken teslim tarihi önce") +
    fld("Günlük iş çeşidi", `<input type="number" min="0" max="20" data-p="gunluk_is_cesidi" value="${A.plan.gunluk_is_cesidi}">`,
        "Bir kaynak gün içinde en fazla kaç farklı ürüne geçebilir. 1 = gün içinde tek iş, 0 = sınırsız. Teslim tarihi gelmiş işler bu kuralı aşabilir.") +
    fld("İş başına azami kaynak", `<input type="number" min="0" max="20" data-p="is_basina_kaynak" value="${A.plan.is_basina_kaynak}">`,
        "Bir iş aynı gün en fazla kaç kişiye/ekibe bölünsün. 1 = işi erken bitirmek için başkalarına bölme, tek kişi birkaç gün çalışsın. 0 = sınırsız. Teslim tarihi gelmiş işler bu kuralı aşabilir.") +
    fld("Erken sevkiyat", `<select data-p="erken_sevkiyat">`+
        ["HAYIR","EVET"].map(x=>`<option ${A.plan.erken_sevkiyat===x?"selected":""}>${x}</option>`).join("")+
        `</select>`,
        "Mal erken hazır olsa bile teslim tarihinden önce sevk edilsin mi? HAYIR = stokta bekler, çıkış teslim haftasında görünür.") +
    fld("Takviye emniyet payı (gün)", `<input type="number" min="0" max="20" data-p="takviye_payi" value="${A.plan.takviye_payi}">`,
        "İş tek kişiyle teslime yetişmiyorsa plan ekibi büyütür. Bu pay kadar erken bitirecek şekilde hesaplar; 0 = tam gününde, 2 = iki iş günü önce.") +
    fld("Minimum kalıp kalış (gün)", `<input type="number" min="0" data-p="min_kalip_kalis" value="${A.plan.min_kalip_kalis}">`,
        "Kalıp bir makineye bağlandıktan sonra en az kaç gün orada kalsın. 0 = serbest") +
    fld("Tatiller", `<input type="text" data-p="tatiller" value="${A.plan.tatiller||""}" placeholder="30.08.2027;29.10.2027">`,
        "Noktalı virgülle ayır") +
    fld("Senaryo analizi", `<select data-p="senaryo">`+
        ["HAYIR","EVET"].map(x=>`<option ${A.plan.senaryo===x?"selected":""}>${x}</option>`).join("")+`</select>`,
        "Darboğaza kademeli kaynak eklenirse ne olur — büyük veride uzun sürer") +
    fld("Satınalma kısıtı", `<select data-p="satinalma">`+
        ["HAYIR","EVET"].map(x=>`<option ${A.plan.satinalma===x?"selected":""}>${x}</option>`).join("")+`</select>`,
        "HAYIR = hammadde her zaman bulunur varsayılır, açık raporlanır");

  const adlar = {kisi:"Kişiler (paketleme, montaj, çapak…)", operator:"Enjeksiyon operatörleri",
                 makine:"Enjeksiyon makineleri", kalip:"Kalıplar", firin:"Fırınlar"};
  el("t-saat").innerHTML = Object.entries(adlar).map(([k,ad])=>
    fld(ad, `<div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
        <input type="number" min="0" max="24" step="0.5" data-v="${k}" style="width:92px"
          value="${A.varsayilan[k].saat}"> <span class="hint">sa/gün</span>
        ${gunKutu(k, A.varsayilan[k].gunler)}</div>`)).join("")+
    `<p class="hint" style="margin-top:12px">Bu değerler tüm kaynaklara varsayılan olarak uygulanır.
      Tek tek değiştirmek için <b>Kaynak bazında</b> sekmesini kullan.</p>`;

  el("t-op").innerHTML =
    fld("Operatör kısıtı", `<select data-o="kisit">
        <option value="1" ${A.operator.kisit?"selected":""}>Evet — operatör sayısı planı sınırlasın</option>
        <option value="0" ${!A.operator.kisit?"selected":""}>Hayır — sadece makine/kalıp sınırlasın</option></select>`,
        "Enjeksiyon makineleri çalışırken operatör havuzundan da kapasite tüketir") +
    fld("Operatör / makine katsayısı", `<input type="number" min="0.05" max="5" step="0.05" data-o="katsayi" value="${A.operator.katsayi}">`,
        "1 = bir operatör bir makine. Bir operatör 3 makineye bakıyorsa 0,33 yaz.") +
    `<p class="hint" style="margin-top:12px">Havuzdaki kişi sayısı PERSONELLISTESI'ndeki
      ENJEKSIYON grubundan otomatik alınır (${(V.gruplar["ENJ-OPERATOR"]||[]).length} kişi).</p>`;

  const satir = k => { const d=KA[k], o=A.kaynaklar[k]||{};
    const deg = o.saat!==undefined||o.paralel!==undefined||o.gunler!==undefined;
    return `<tr data-k="${k}" class="${deg?"deg":""}">
      <td><b>${k}</b><br><span class="rn">${d.ad||""}</span></td>
      <td><span class="pill">${TIPAD[d.tip]||d.tip}</span></td>
      <td><input type="number" min="0" max="24" step="0.5" data-kk="saat" value="${o.saat??d.saat}"></td>
      <td><input type="number" min="1" data-kk="paralel" value="${o.paralel??d.paralel}"></td>
      <td>${gunKutu("K:"+k, o.gunler??d.gunler)}</td></tr>`; };
  el("t-kay").innerHTML = `<div class="bar" style="position:static;margin-bottom:10px">
      <input id="kayAra" type="search" placeholder="Kaynak ara…" style="min-width:220px">
      <select id="kayTip"><option value="">Tüm tipler</option>${
        [...new Set(Object.values(KA).map(d=>d.tip))].sort().map(t=>`<option value="${t}">${TIPAD[t]||t}</option>`).join("")}</select>
      <span class="hint">Değiştirilen satırlar renklenir; yalnızca onlar kaydedilir.</span></div>
    <div class="kaylist"><table class="kay"><thead><tr><th>Kaynak</th><th>Tip</th>
      <th>Saat/gün</th><th>Paralel</th><th>Çalışılan günler</th></tr></thead>
      <tbody id="kayBody">${Object.keys(KA).sort().map(satir).join("")}</tbody></table></div>`;
  const sz = ()=>{ const q=el("kayAra").value.toLocaleLowerCase("tr"), t=el("kayTip").value;
    el("kayBody").querySelectorAll("tr").forEach(tr=>{ const k=tr.dataset.k;
      tr.style.display = (!t||KA[k].tip===t) &&
        (!q||(k+" "+(KA[k].ad||"")).toLocaleLowerCase("tr").includes(q)) ? "" : "none"; }); };
  el("kayAra").oninput = el("kayTip").onchange = sz;
  ayarDinle();
}

function ayarDinle(){
  el("ayar").querySelectorAll("[data-p]").forEach(i=>i.onchange=()=>{
    const k=i.dataset.p; A.plan[k] = i.type==="number" ? sayi(i.value,0) : i.value; bozuldu(); });
  el("ayar").querySelectorAll("[data-v]").forEach(i=>i.onchange=()=>{
    A.varsayilan[i.dataset.v].saat = sayi(i.value,0); bozuldu(); });
  el("ayar").querySelectorAll("[data-o]").forEach(i=>i.onchange=()=>{
    A.operator[i.dataset.o] = i.dataset.o==="kisit" ? i.value==="1" : sayi(i.value,1); bozuldu(); });
  el("ayar").querySelectorAll("[data-gun]").forEach(i=>i.onchange=()=>{
    const ad=i.dataset.gun, ix=+i.dataset.i;
    if(ad.startsWith("K:")){ const k=ad.slice(2);
      const g=(A.kaynaklar[k]?.gunler)??[...KA[k].gunler]; g[ix]=i.checked?1:0;
      (A.kaynaklar[k] ||= {}).gunler=g; el("kayBody").querySelector(`tr[data-k="${CSS.escape(k)}"]`)?.classList.add("deg");
    } else { A.varsayilan[ad].gunler[ix]=i.checked?1:0; }
    bozuldu(); });
  el("ayar").querySelectorAll("[data-kk]").forEach(i=>i.onchange=()=>{
    const tr=i.closest("tr"), k=tr.dataset.k;
    (A.kaynaklar[k] ||= {})[i.dataset.kk] = sayi(i.value,0);
    tr.classList.add("deg"); bozuldu(); });
}
function bozuldu(){ el("ayarDurum").textContent="kaydedilmedi"; el("ayarDurum").className="durum err"; }

function ayarPaket(){
  const k = {};
  for(const [kod,o] of Object.entries(A.kaynaklar)){
    const d=KA[kod]||{}, t={};
    if(o.saat!==undefined && o.saat!==d.saat) t.saat=o.saat;
    if(o.paralel!==undefined && o.paralel!==d.paralel) t.paralel=o.paralel;
    if(o.gunler && JSON.stringify(o.gunler)!==JSON.stringify(d.gunler)) t.gunler=o.gunler;
    if(Object.keys(t).length) k[kod]=t;
  }
  return {kaydedildi:new Date().toLocaleString("tr-TR"), plan:A.plan,
          varsayilan:A.varsayilan, operator:A.operator, kaynaklar:k};
}
async function dosyaKaydet(ad, nesne, durumEl){
  const metin = JSON.stringify(nesne, null, 2);
  try{
    if(window.showSaveFilePicker){
      const h = await showSaveFilePicker({suggestedName:ad,
        types:[{description:"JSON", accept:{"application/json":[".json"]}}]});
      const w = await h.createWritable(); await w.write(metin); await w.close();
      durumEl.textContent = ad+" kaydedildi"; durumEl.className="durum ok"; return;
    }
  }catch(e){ if(e.name==="AbortError"){ return; } }
  const a=document.createElement("a");
  a.href=URL.createObjectURL(new Blob([metin],{type:"application/json"}));
  a.download=ad; a.click(); URL.revokeObjectURL(a.href);
  durumEl.textContent = ad+" indirildi — program klasörüne taşı";
  durumEl.className="durum ok";
}

el("ayarAc").onclick = ()=>{ ayarCiz(); el("ayar").hidden=false;
  el("ayarDurum").textContent=""; el("ayarDurum").className="durum"; };
el("ayarKapat").onclick = ()=> el("ayar").hidden=true;
el("ayar").onclick = e=>{ if(e.target.id==="ayar") el("ayar").hidden=true; };
el("ayarTabs").onclick = e=>{ const t=e.target.dataset?.t; if(!t) return;
  el("ayarTabs").querySelectorAll("button").forEach(b=>b.setAttribute("aria-selected", b.dataset.t===t));
  ["plan","saat","op","kay"].forEach(x=> el("t-"+x).hidden = x!==t); };
el("ayarKaydet").onclick = ()=> dosyaKaydet("ayarlar.json", ayarPaket(), el("ayarDurum"));
el("ayarIndir").onclick = ()=>{ const a=document.createElement("a");
  a.href=URL.createObjectURL(new Blob([JSON.stringify(ayarPaket(),null,2)],{type:"application/json"}));
  a.download="ayarlar.json"; a.click(); URL.revokeObjectURL(a.href); };
el("ayarSifirla").onclick = ()=>{ A=JSON.parse(JSON.stringify(A0)); A.kaynaklar={}; ayarCiz();
  el("ayarDurum").textContent="varsayılanlara dönüldü — kaydetmeyi unutma"; el("ayarDurum").className="durum err"; };

// ===================== IS EMRI -> KISI ATAMA =====================
// Grup havuzuna atanmis isler (orn. ENJ-OPERATOR) ve birden fazla alternatifi
// olan isler icin planlamaci gercekte kimin yapacagini secer. Secimler
// atamalar.json'a kaydedilir; plani degistirmez, sevk karari olarak saklanir.
let ATAMA = Object.assign({}, V.atamalar || {});
const kisiAdi = k => { const i=KYN.findIndex(x=>x[0]===k);
  return i>=0 && KYN[i][1] ? KYN[i][1].split(" - ")[0] : ""; };

function adaylar(e){
  const set = new Map();
  if(e[11] && V.gruplar[e[11]]) for(const k of V.gruplar[e[11]]) set.set(k, "grup");
  if(e[10] && V.gruplar[e[10]]) for(const k of V.gruplar[e[10]]) set.set(k, "havuz");
  const alt = (V.rota[MAL[e[1]][0]]||[]).slice().sort((a,b)=>a[0]-b[0]);
  if(alt.length > 1) for(const [,k] of alt) if(!set.has(k)) set.set(k, "alternatif");
  return [...set.keys()];
}
function kisiSecici(e){
  if(e[12]>1){
    // Is ayni anda birden fazla kisiyi istiyor: tek kisi secilemez,
    // motorun kurdugu ekip gosterilir.
    const ek = [e[2]].concat(e[13]||[]);
    return `<span class="rn">${ek.map(i=>kisiAdi(KYN[i][0])||KYN[i][0]).join(" · ")}</span>
            <br><span class="pill warn">birlikte çalışır</span>`;
  }
  const ad = adaylar(e);
  if(!ad.length) return `<span class="rn">${KYN[e[2]][1]?KYN[e[2]][1].split(" - ")[0]:KYN[e[2]][0]}</span>`;
  const sec = ATAMA[e[0]] || "";
  return `<select class="kisi${sec?" set":""}" data-em="${e[0]}">
    <option value="">— plan kaynağı —</option>`+
    ad.map(k=>`<option value="${k}" ${sec===k?"selected":""}>${k}${kisiAdi(k)?" · "+kisiAdi(k):""}</option>`).join("")+
    `</select>`;
}
function ataDinle(){
  el("isler").querySelectorAll("select.kisi").forEach(sl=>{
    sl.onchange = ()=>{
      if(sl.value) ATAMA[sl.dataset.em] = sl.value; else delete ATAMA[sl.dataset.em];
      sl.classList.toggle("set", !!sl.value);
      el("ataDurum").textContent = Object.keys(ATAMA).length+" atama · kaydedilmedi";
      el("ataDurum").className = "durum err";
    };
  });
}


// ================= SIPARIS AKISI (urun agaci bazli) =================
// Soldan saga: hammadde -> ara urun -> mamul -> siparis. Her kutuda o haftanin
// acilis stogu, uretimi, cikisi ve kapanis stogu var; ok uzerindeki sayi o hafta
// ust malzeme icin tuketilen miktardir.
const AKIS = (V.akis) || {hafta:{}, seviye:{}, tip:{}};
let akisSec = null, akisZoom = 0, akisMod = "sip", akisQ = "";
const AKW = 206, AKH = 76, AKGX = 86, AKGY = 10;

function akisTip(mi){
  const k = MAL[mi] ? MAL[mi][0] : "";
  if((AKIS.tip||{})[mi] === "S") return "HAM";
  const m = /^[A-Z]+/.exec(k);
  return m ? m[0].slice(0,4) : "—";
}
function kis(s,n){ s = s||""; return s.length>n ? s.slice(0,n-1)+"…" : s; }
// Aciklamayi iki satira kelime sinirindan boler
function iki(s,n){
  s = (s||"").trim();
  if(s.length<=n) return [s,""];
  let k = s.lastIndexOf(" ", n);
  if(k < n*0.5) k = n;
  return [s.slice(0,k), kis(s.slice(k).trim(), n)];
}
function esc(s){ return String(s==null?"":s).replace(/[&<>]/g, c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c])); }

function cizAkis(hb){
  const kutu = el("akis");
  const A = AKIS.hafta[hb];
  if(!A || (!A.n.length && !A.s.length)){
    kutu.innerHTML = `<h2>Sipariş akışı</h2>
      <p class="note">Bu haftada stok hareketi ya da teslim tarihi yok.</p>`;
    return;
  }
  // ---- dugumler
  const SEV = AKIS.seviye || {};
  const D = new Map();                                   // anahtar -> dugum
  let maxS = 0;
  A.n.forEach(d => maxS = Math.max(maxS, SEV[d[0]]||0));
  A.n.forEach(d => D.set("m"+d[0], {k:"m"+d[0], t:"m", mi:d[0],
      ac:d[1], ur:d[2], tu:d[3], sv:d[4], kap:d[5], ie:d[6], sa:d[7],
      gir:(d[8]==null?d[2]:d[8]),
      col: maxS - (SEV[d[0]]||0)}));
  A.s.forEach((s,i) => D.set("s"+i, {k:"s"+i, t:"s", mi:s[2], belge:s[0],
      kalem:s[1], miktar:s[3], teslim:s[4], durum:s[5], bitis:s[6], gec:s[7],
      col: maxS+1}));

  // ---- kenarlar: alt (sol) -> ust (sag)
  const K = [];
  A.e.forEach(e => { const a="m"+e[1], b="m"+e[0];
    if(D.has(a) && D.has(b)) K.push({a,b,q:e[2],t:"u"}); });
  A.s.forEach((s,i) => { const a="m"+s[2];
    if(D.has(a)) K.push({a, b:"s"+i, q:s[3], t:"s", gec:s[7]>0}); });

  // ---- kapsam: siparislerden geriye dogru kapanis (tum zincir)
  let gorunur = null;
  if(akisMod !== "tum"){
    const geri = new Map();
    K.forEach(k => { if(!geri.has(k.b)) geri.set(k.b,[]); geri.get(k.b).push(k.a); });
    gorunur = new Set();
    const sira = [...D.values()]
      .filter(n => n.t==="s" && (akisMod==="sip" || n.gec>0)).map(n=>n.k);
    sira.forEach(x=>gorunur.add(x));
    while(sira.length){ const x = sira.pop();
      (geri.get(x)||[]).forEach(y => { if(!gorunur.has(y)){ gorunur.add(y); sira.push(y); } }); }
  }
  const dugumler = [...D.values()].filter(n => !gorunur || gorunur.has(n.k));
  const anah = new Set(dugumler.map(n=>n.k));
  const kenarlar = K.filter(k => anah.has(k.a) && anah.has(k.b));
  if(!dugumler.length){
    kutu.innerHTML = `<h2>Sipariş akışı</h2>
      <p class="note">Bu kapsamda gösterilecek akış yok — kapsamı değiştirin.</p>`
      + akisBar(0, 0, A.s.filter(x=>x[7]>0).length, A.s.length);
    akisBarDinle(hb); return;
  }

  // ---- bos sutunlari kaldir (filtreden sonra sutunlar seyrek kalabiliyor)
  const kullanilan = [...new Set(dugumler.map(n=>n.col))].sort((a,b)=>a-b);
  const yeniKol = new Map(kullanilan.map((c,i)=>[c,i]));
  dugumler.forEach(n => n.col = yeniKol.get(n.col));

  // ---- sutun icinde sira: baryantrik (cizgi kesismesini azaltir)
  const kol = [];
  dugumler.forEach(n => { (kol[n.col] = kol[n.col] || []).push(n); });
  kol.forEach(c => c && c.sort((a,b)=>
    (MAL[a.mi]?MAL[a.mi][0]:"").localeCompare(MAL[b.mi]?MAL[b.mi][0]:"")));
  const ileri = new Map(), geri = new Map();
  kenarlar.forEach(k => {
    if(!ileri.has(k.a)) ileri.set(k.a,[]); ileri.get(k.a).push(k.b);
    if(!geri.has(k.b))  geri.set(k.b,[]);  geri.get(k.b).push(k.a);
  });
  const yer = new Map();
  const yaz = ()=> kol.forEach(c => c && c.forEach((n,i)=>yer.set(n.k,i)));
  yaz();
  const ortala = (n, harita) => {
    const ks = harita.get(n.k)||[];
    const v = ks.map(x=>yer.get(x)).filter(x=>x!=null);
    return v.length ? v.reduce((a,b)=>a+b,0)/v.length : yer.get(n.k);
  };
  for(let p=0;p<4;p++){
    for(let c=1;c<kol.length;c++) if(kol[c]){
      kol[c].sort((a,b)=>ortala(a,geri)-ortala(b,geri)); yaz(); }
    for(let c=kol.length-2;c>=0;c--) if(kol[c]){
      kol[c].sort((a,b)=>ortala(a,ileri)-ortala(b,ileri)); yaz(); }
  }
  // ---- ayni urun ailesini alt alta topla
  // Kodlardaki "012-00012" gibi ortak govde ayni urunun seviyelerini isaret
  // eder (ENJ012-00012 / BLZ012-00012 / MAM012-00012 ...). Once kesisme
  // azaltildi; simdi ayni aile sutun icinde bitisik ve sutunlar arasinda
  // benzer yukseklikte duracak sekilde yeniden siralaniyor.
  // Kodun iki parcasi: SERI (MAM012 / BLZ012 / ENJ012 -> "012") ve
  // AILE (seri + parca no -> "012-00012"). Once ayni seri, sonra ayni aile
  // alt alta gelir; boylece bir urunun butun seviyeleri bitisik ve sutunlar
  // arasinda ayni yukseklikte durur.
  const kodu = n => MAL[n.mi] ? MAL[n.mi][0] : "";
  const parcala = n => {
    const m = /(\d{3})-(\d{4,6})/.exec(kodu(n));
    return m ? [m[1], m[1] + "-" + m[2]] : ["~" + kodu(n), "~" + kodu(n)];
  };
  const ortKonum = anahtarla => {
    const t = new Map();
    dugumler.forEach(n=>{
      const c = kol[n.col] || [], i = yer.get(n.k) || 0;
      const p = c.length > 1 ? i/(c.length-1) : 0.5;
      const a = anahtarla(n);
      const v = t.get(a) || [0,0];
      v[0] += p; v[1] += 1; t.set(a, v);
    });
    return new Map([...t.entries()]
      .map(([a,v]) => [a, v[0]/v[1]])
      .sort((x,y) => x[1]-y[1] || String(x[0]).localeCompare(String(y[0])))
      .map(([a], i) => [a, i]));
  };
  const seriSira = ortKonum(n => parcala(n)[0]);
  const aileSira = ortKonum(n => parcala(n)[1]);
  kol.forEach(c=>{ if(!c) return;
    c.sort((x,y)=>{
      const [sx, ax] = parcala(x), [sy, ay] = parcala(y);
      return (seriSira.get(sx) - seriSira.get(sy))
          || (aileSira.get(ax) - aileSira.get(ay))
          || kodu(x).localeCompare(kodu(y))
          || (yer.get(x.k) - yer.get(y.k));
    }); });
  yaz();

  let enUzun = 0;
  for(const c of kol) if(c) enUzun = Math.max(enUzun, c.length);
  const Hsvg = enUzun*(AKH+AKGY) + 46;
  kol.forEach((c,ci)=>{ if(!c) return;
    c.forEach((n,i)=>{ n.x = ci*(AKW+AKGX)+14; n.y = 38 + i*(AKH+AKGY); });
  });
  const Wsvg = kol.length*(AKW+AKGX)+14;

  // ---- vurgulama: secilen dugumun tum zinciri
  const zincir = new Set();
  if(akisSec && anah.has(akisSec)){
    // ileri ve geri yuruyus ayri 'gorulen' kumesiyle yapilmali; ortak kume
    // kullanilirsa ikinci yuruyus baslangic dugumunde hemen duruyor
    const gez = (bas, harita) => { const gor=new Set(), s=[bas];
      while(s.length){ const x=s.pop(); if(gor.has(x)) continue; gor.add(x); zincir.add(x);
        (harita.get(x)||[]).forEach(y=>s.push(y)); } };
    gez(akisSec, ileri); gez(akisSec, geri);
  }
  const aramaVar = akisQ.length>0;
  const uyan = n => !aramaVar ? false :
    ((MAL[n.mi]?MAL[n.mi][0]+" "+MAL[n.mi][1]:"")+" "+(n.belge||""))
      .toLocaleLowerCase("tr").includes(akisQ);
  const eslesen = aramaVar ? dugumler.filter(uyan) : [];

  // ---- ciz
  const maxQ = Math.max(1, ...kenarlar.map(k=>k.q));
  const kalinlik = q => 0.9 + 2.4*Math.log10(1+9*Math.max(0,q)/maxQ);
  const etiketGoster = kenarlar.length <= 48;
  let sv = `<svg viewBox="0 0 ${Wsvg} ${Hsvg}" width="${Math.round(Wsvg*akisZoom)}"
      role="img" aria-label="Haftalık malzeme akışı: hammaddeden siparişe">
    <defs>
      <marker id="akok" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="7"
        markerHeight="7" orient="auto-start-reverse">
        <path d="M0 0 L8 4 L0 8 z" fill="var(--axis)"/></marker>
      <marker id="akokg" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="7"
        markerHeight="7" orient="auto-start-reverse">
        <path d="M0 0 L8 4 L0 8 z" fill="var(--good)"/></marker>
      <marker id="akokk" viewBox="0 0 8 8" refX="7" refY="4" markerWidth="7"
        markerHeight="7" orient="auto-start-reverse">
        <path d="M0 0 L8 4 L0 8 z" fill="var(--critical)"/></marker>
    </defs>`;
  // sutun basliklari
  kol.forEach((c,ci)=>{ if(!c||!c.length) return;
    let ad;
    if(c.every(n=>n.t==="s")) ad = "SİPARİŞ";
    else {
      const say = new Map();
      c.forEach(n=>{ const t=akisTip(n.mi); say.set(t,(say.get(t)||0)+1); });
      ad = [...say.entries()].sort((a,b)=>b[1]-a[1]).slice(0,3)
             .map(x=>x[0]).join(" · ");
    }
    sv += `<text class="ak-bas" x="${ci*(AKW+AKGX)+14}" y="22">${ad} <tspan
      fill="var(--muted)" font-weight="400">(${c.length})</tspan></text>`;
  });
  // kenarlar
  kenarlar.forEach(k=>{
    const a=D.get(k.a), b=D.get(k.b);
    const x1=a.x+AKW, y1=a.y+AKH/2, x2=b.x, y2=b.y+AKH/2, mx=(x1+x2)/2;
    const v = zincir.has(k.a) && zincir.has(k.b);
    const mk = k.t==="s" ? (k.gec?"akokk":"akokg") : "akok";
    sv += `<path class="ak-k ${k.t==="s"?"sip":""} ${k.gec?"gec":""} ${v?"vurgu":""}"
      d="M${x1} ${y1} C${mx} ${y1} ${mx} ${y2} ${x2-9} ${y2}"
      stroke-width="${kalinlik(k.q).toFixed(2)}" marker-end="url(#${mk})"/>`;
    if(etiketGoster || v)
      sv += `<text class="ak-et ${v?"vurgu":""}" x="${mx}" y="${(y1+y2)/2-4}"
        text-anchor="middle">${fmt(k.q)}</text>`;
  });
  // dugumler
  dugumler.forEach(n=>{
    const v = zincir.has(n.k) || uyan(n);
    if(n.t==="s"){
      const gec = n.gec>0;
      sv += `<g class="ak-n sip ${gec?"gec":""} ${v?"vurgu":""} ${uyan(n)?"bulundu":""}" data-k="${n.k}">
        <rect class="gov" x="${n.x}" y="${n.y}" width="${AKW}" height="${AKH}" rx="9"/>
        <text x="${n.x+12}" y="${n.y+19}">${esc(n.belge)} / ${esc(n.kalem)}</text>
        <text class="ad" x="${n.x+12}" y="${n.y+32}">${esc(kis(MAL[n.mi]?MAL[n.mi][0]:"",30))}</text>
        <text class="ad" x="${n.x+12}" y="${n.y+42}">${esc(iki(MAL[n.mi]?MAL[n.mi][1]:"",36)[0])}</text>
        <text class="et" x="${n.x+12}" y="${n.y+56}">TESLİM</text>
        <text class="et" x="${n.x+96}" y="${n.y+56}">MİKTAR</text>
        <text class="et" x="${n.x+156}" y="${n.y+56}">DURUM</text>
        <text class="sy" x="${n.x+12}" y="${n.y+69}">${trt(d2(n.teslim))}</text>
        <text class="sy" x="${n.x+96}" y="${n.y+69}">${fmt(n.miktar)}</text>
        <text class="sy ${gec?"kot":""}" x="${n.x+156}" y="${n.y+69}">${gec?("+"+n.gec+" gün"):"zamanında"}</text>
      </g>`;
      return;
    }
    const sinif = n.kap < -0.001 ? "eksi" : (n.ur>0 ? "uretim" : "");
    sv += `<g class="ak-n ${sinif} ${v?"vurgu":""} ${uyan(n)?"bulundu":""}" data-k="${n.k}">
      <rect class="gov" x="${n.x}" y="${n.y}" width="${AKW}" height="${AKH}" rx="9"/>
      <text x="${n.x+12}" y="${n.y+18}">${esc(MAL[n.mi]?MAL[n.mi][0]:"")}</text>
      <text class="et" x="${n.x+AKW-12}" y="${n.y+18}" text-anchor="end">${akisTip(n.mi)}${n.ie?" · "+n.ie+" İE":""}</text>
      <text class="ad" x="${n.x+12}" y="${n.y+31}">${esc(iki(MAL[n.mi]?MAL[n.mi][1]:"",36)[0])}</text>
      <text class="ad" x="${n.x+12}" y="${n.y+41}">${esc(iki(MAL[n.mi]?MAL[n.mi][1]:"",36)[1])}</text>
      <text class="et" x="${n.x+12}"  y="${n.y+56}">AÇILIŞ</text>
      <text class="et" x="${n.x+67}"  y="${n.y+56}">ÜRETİM</text>
      <text class="et" x="${n.x+124}" y="${n.y+56}">ÇIKIŞ</text>
      <text class="et" x="${n.x+172}" y="${n.y+56}">KAPANIŞ</text>
      <text class="sy" x="${n.x+12}"  y="${n.y+69}">${fmt(n.ac)}</text>
      <text class="sy art" x="${n.x+67}"  y="${n.y+69}">${n.ur>0?"+"+fmt(n.ur):"–"}</text>
      <text class="sy eks" x="${n.x+124}" y="${n.y+69}">${(n.tu+n.sv)>0?"−"+fmt(n.tu+n.sv):"–"}</text>
      <text class="sy ${n.kap<-0.001?"kot":""}" x="${n.x+172}" y="${n.y+69}">${fmt(n.kap)}</text>
    </g>`;
  });
  sv += `</svg>`;

  const gecSay = A.s.filter(x=>x[7]>0).length;
  kutu.innerHTML = `<h2>Sipariş akışı — ürün ağacı üzerinde</h2>
    <p class="note">Soldan sağa malzeme akışı: en alt seviyeden mamule, mamulden siparişe.
      Her kutuda bu haftanın <b>açılış stoğu</b>, <b>üretimi</b> (açılan iş emirleri),
      <b>çıkışı</b> (üst seviyeye tüketim + sevkiyat) ve <b>kapanış stoğu</b> var;
      ok üzerindeki sayı üst malzeme için o hafta tüketilen miktardır.
      Bir kutuya tıklayınca o malzemenin zinciri uçtan uca kalır.</p>`
    + akisBar(dugumler.length, kenarlar.length, gecSay, A.s.length, eslesen.length)
    + `<div class="akis-kutu">${sv}</div>`
    + `<div class="legend">
         <span><b>İE</b> = o hafta o malzeme için açılan iş emri sayısı</span>
         <span style="margin-left:10px">·</span>
         <span>Mavi kenarlık = bu hafta üretiliyor</span>
         <span style="margin-left:10px">·</span>
         <span>Kırmızı kenarlık = kapanış stoğu eksi ya da sipariş gecikmeli</span>
       </div>`;
  akisBarDinle(hb);
  // zoom verilmemisse genisligi kutuya sigdir
  const sarmal = kutu.querySelector(".akis-kutu");
  const svgEl = sarmal.querySelector("svg");
  // Genislige sigdirirken 0.8'in altina inme: daha kucugunde yazilar okunmuyor,
  // bunun yerine kutu yatay kaydirilir.
  const oran = akisZoom || Math.min(1, Math.max(0.8,
                 (sarmal.clientWidth - 18) / Wsvg));
  svgEl.setAttribute("width", Math.round(Wsvg*oran));
  kutu.querySelectorAll(".ak-n").forEach(g=>{
    g.onclick = ()=>{ akisSec = (akisSec===g.dataset.k) ? null : g.dataset.k; cizAkis(hb); };
    g.onmousemove = e => tipGoster(e, akisTip2(D.get(g.dataset.k)));
    g.onmouseleave = tipKapat;
  });
  kutu.querySelector(".akis-kutu").classList.toggle("akis-sonuk", !!akisSec || aramaVar);
  // arama varsa ilk eslesmeyi gorus alanina getir
  const ilk = kutu.querySelector(".ak-n.bulundu");
  if(ilk){
    const r = ilk.getBoundingClientRect(), k = sarmal.getBoundingClientRect();
    sarmal.scrollLeft += (r.left - k.left) - sarmal.clientWidth/2 + r.width/2;
    sarmal.scrollTop  += (r.top  - k.top)  - sarmal.clientHeight/2 + r.height/2;
  }
}

function akisTip2(n){
  if(!n) return "";
  const ad = MAL[n.mi] ? `<b>${esc(MAL[n.mi][0])}</b><br><span class="m">${esc(MAL[n.mi][1]||"")}</span>` : "";
  if(n.t==="s") return `<b>${esc(n.belge)} / ${esc(n.kalem)}</b><hr>${ad}
    <div class="r"><span class="m">Miktar</span><span>${fmt(n.miktar)}</span></div>
    <div class="r"><span class="m">Teslim</span><span>${n.teslim}</span></div>
    <div class="r"><span class="m">Planlanan bitiş</span><span>${n.bitis||"–"}</span></div>
    <div class="r"><span class="m">Durum</span><span>${n.gec>0?("+"+n.gec+" gün gecikme"):"zamanında"}</span></div>`;
  const fark = Math.abs(n.gir - n.ur) > 0.5;
  return `${ad}<hr>
    <div class="r"><span class="m">Açılış stoğu</span><span>${fmt(n.ac)}</span></div>
    <div class="r"><span class="m">Bu hafta üretim</span><span>${fmt(n.ur)}</span></div>
    ${fark?`<div class="r"><span class="m">Stoğa giren</span><span>${fmt(n.gir)}</span></div>`:""}
    <div class="r"><span class="m">Üst seviyeye tüketim</span><span>${fmt(n.tu)}</span></div>
    <div class="r"><span class="m">Sevkiyat</span><span>${fmt(n.sv)}</span></div>
    <div class="r"><span class="m">Kapanış stoğu</span><span>${fmt(n.kap)}</span></div>
    <hr><div class="r"><span class="m">Stok denkliği</span>
      <span>${fmt(n.ac)} + ${fmt(n.gir)} − ${fmt(n.tu+n.sv)} = ${fmt(n.ac+n.gir-n.tu-n.sv)}</span></div>
    ${fark?`<div class="r"><span class="m" style="font-size:11px">Üretim ertesi gün stoğa girdiği için
       "üretim" ile "stoğa giren" hafta sonunda farklı olabilir.</span><span></span></div>`:""}
    <div class="r"><span class="m">İş emri</span><span>${n.ie} adet · ${fmt(n.sa,1)} saat</span></div>`;
}

function akisBar(nd, nk, gec, sip, bulunan){
  return `<div class="akis-bar">
    <select id="akisMod">
      <option value="sip" ${akisMod==="sip"?"selected":""}>bu haftaki siparişlere bağlı zincir${sip?" ("+sip+" sipariş)":""}</option>
      <option value="gec" ${akisMod==="gec"?"selected":""}>sadece geciken siparişler${gec?" ("+gec+")":""}</option>
      <option value="tum" ${akisMod==="tum"?"selected":""}>bu haftaki tüm hareketler</option>
    </select>
    <input type="search" id="akisAra" placeholder="Malzeme veya sipariş ara…" value="${esc(akisQ)}">
    <button id="akisEksi">−</button><button id="akisArti">+</button>
    <button id="akisSigdir">sığdır</button>
    <button id="akisSifirla">sıfırla</button>
    <span class="spacer"></span>
    <span class="hint">${bulunan!=null && akisQ ? `<b>${bulunan} eşleşme</b> · ` : ""}${nd} kutu · ${nk} bağlantı</span>
  </div>`;
}
function akisBarDinle(hb){
  const a = el("akisAra");
  if(a){ a.oninput = ()=>{ akisQ = a.value.trim().toLocaleLowerCase("tr");
         cizAkis(hb); el("akisAra").focus(); }; }
  const g = el("akisMod");
  if(g) g.onchange = ()=>{ akisMod = g.value; akisSec = null; akisZoom = 0; cizAkis(hb); };
  const z = ()=> akisZoom || 1;
  if(el("akisEksi")) el("akisEksi").onclick = ()=>{ akisZoom=Math.max(.4,z()-.15); cizAkis(hb); };
  if(el("akisArti")) el("akisArti").onclick = ()=>{ akisZoom=Math.min(2,z()+.15); cizAkis(hb); };
  if(el("akisSigdir")) el("akisSigdir").onclick = ()=>{ akisZoom=0; cizAkis(hb); };
  if(el("akisSifirla")) el("akisSifirla").onclick = ()=>{
    akisZoom=0; akisSec=null; akisQ=""; akisMod="sip"; cizAkis(hb); };
}


// ============ KAYNAK CIZELGESI: hangi kaynakta hangi is emri calisiyor ============
// Satir = kaynak (kisi / kalip / makine / firin), sutun = haftanin gunu.
// Bir is emri basladigi gunden bittigi gune kadar bir cubuk olarak cizilir;
// ayni kaynakta cakisan isler alt alta istiflenir.
let kcSec = null;

function cizKaynakCizelge(hb, gunler, em){
  const kutu = el("kcizelge");
  const q = el("ara").value.trim().toLocaleLowerCase("tr");
  const tipF = selT.value;
  const gIx = g => gunler.indexOf(g);

  // kaynak -> o kaynakta calisan is emirleri (kalip, makine ve ekip uyeleri dahil)
  const sat = new Map();
  const koy = (ki, e, rol) => {
    if(ki == null || ki < 0) return;
    if(tipF && KYN[ki][2] !== tipF) return;
    if(q && !((KYN[ki][0]+" "+KYN[ki][1]+" "+MAL[e[1]][0]+" "+MAL[e[1]][1]+" "+e[0])
              .toLocaleLowerCase("tr").includes(q))) return;
    if(!sat.has(ki)) sat.set(ki, []);
    sat.get(ki).push({e, rol});
  };
  em.forEach(e=>{
    koy(e[2], e, KYN[e[2]][2]==="KALIP" ? "kalip" : "ana");
    if(e[3] >= 0) koy(e[3], e, "makine");
    (e[13]||[]).forEach(k => koy(k, e, "ekip"));
  });

  if(!sat.size){
    kutu.innerHTML = `<h2>Kaynak çizelgesi</h2>
      <p class="note">Bu haftada (ve bu filtrede) iş emri yok.</p>`;
    return;
  }

  // satirlari tipe ve koda gore sirala
  const tipSira = {KISI:0, KALIP:1, MAKINE:2, FIRIN:3};
  const satirlar = [...sat.entries()].sort((a,b)=>{
    const ta = tipSira[KYN[a[0]][2]] ?? 9, tb = tipSira[KYN[b[0]][2]] ?? 9;
    return ta-tb || KYN[a[0]][0].localeCompare(KYN[b[0]][0]);
  });

  // gun basliklari
  let gh = "";
  gunler.forEach((g,i)=>{
    const d = d2(g), cal = satirlar.some(([ki]) => kapasite(ki,g) > 0);
    gh += `<span class="${cal?"":"tatil"}">${["PZT","SAL","ÇAR","PER","CUM","CMT","PAZ"][i]}
      ${trt(d)}</span>`;
  });

  let govde = "";
  let toplamIs = 0;
  for(const [ki, lst] of satirlar){
    // cakismayan isleri ayni alt satira yerlestir (istifleme)
    lst.sort((a,b)=> gIx(a.e[4]) - gIx(b.e[4]) || gIx(b.e[5]) - gIx(a.e[5]));
    const son = [];                       // alt satir -> son dolu gun
    const yer = [];
    lst.forEach(x=>{
      const a = Math.max(0, gIx(x.e[4]) < 0 ? 0 : gIx(x.e[4]));
      const b = Math.min(6, gIx(x.e[5]) < 0 ? 6 : gIx(x.e[5]));
      let r = 0;
      while(son[r] != null && son[r] >= a) r++;
      son[r] = b;
      yer.push({...x, a, b, r});
    });
    const nr = son.length;
    toplamIs += lst.length;
    const bar = yer.map(x=>{
      const e = x.e, gen = (x.b - x.a + 1);
      const belgeler = String(e[9]||"").split(",")
            .map(x=>x.trim().split(" ")[0].split("/")[0]);
      const gec = (V.geciken||[]).some(z => belgeler.includes(String(z.belge)));
      // sira: kod -> miktar -> aciklama -> is emri. Dar cubukta once aciklama,
      // sonra is emri kirpilir; kod ve miktar her zaman okunur kalir.
      const etiket = `<b>${MAL[e[1]][0]}</b>`
                   + `<i class="mk">${fmt(e[7])}</i>`
                   + `<em>${MAL[e[1]][1]||""}</em>`
                   + `<i class="no">${e[0]}</i>`;
      const baslik = `${MAL[e[1]][0]} — ${MAL[e[1]][1]||""}\n`
        + `${fmt(e[7])} adet · ${fmt(e[8],2)} saat · ${e[4]} → ${e[5]}\n${e[0]}`;
      return `<div class="kc-bar ${x.rol} ${gec?"gec":""} ${kcSec===e[0]+"|"+ki?"sec":""}"
        title="${baslik.replace(/"/g,"&quot;")}" data-em="${e[0]}" data-ki="${ki}"
        style="left:calc(${x.a}*100%/7 + 3px);width:calc(${gen}*100%/7 - 6px);
               top:${6 + x.r*36}px">${etiket}</div>`;
    }).join("");
    govde += `<div class="kc-ad"><b>${KYN[ki][0]}</b>
        <span>${(KYN[ki][1]||"").split(" - ")[0] || TIPAD[KYN[ki][2]]||KYN[ki][2]}</span></div>
      <div class="kc-trk" style="height:${nr*36 + 12}px">${bar}</div>`;
  }

  kutu.innerHTML = `<h2>Kaynak çizelgesi — hangi kaynakta hangi iş emri</h2>
    <p class="note">Her satır bir kaynak, her sütun haftanın bir günü. Çubuk bir iş emrinin
      başlangıcından bitişine kadar uzanır; aynı kaynakta üst üste binen işler alt alta
      istiflenir. <b>Kesik çerçeve</b> = kişi o işte ekibin üyesi, <b>mavi</b> = kalıp/makine,
      <b>kırmızı sol şerit</b> = geciken bir siparişi besleyen iş. Çubuğa tıklayınca
      ayrıntısı açılır; üzerine gelince özeti çıkar.</p>
    <div class="kc-sar"><div class="kc">
      <div class="kc-bas">Kaynak (${satirlar.length})</div>
      <div class="kc-gunler">${gh}</div>
      ${govde}
    </div></div>
    <div class="legend"><span>${toplamIs} çubuk · ${em.length} iş emri</span></div>
    <div id="kcDetay"></div>`;

  kutu.querySelectorAll(".kc-bar").forEach(b=>{
    b.onclick = ()=>{ kcSec = (kcSec === b.dataset.em+"|"+b.dataset.ki)
                              ? null : b.dataset.em+"|"+b.dataset.ki;
                      cizKaynakCizelge(hb, gunler, em);
  cizKapatma(); };
    b.onmousemove = ev => {
      const e = em.find(x=>x[0]===b.dataset.em);
      if(e) tipGoster(ev, kcTip(e, +b.dataset.ki));
    };
    b.onmouseleave = tipKapat;
  });
  if(kcSec){
    const [no, ki] = kcSec.split("|");
    const e = em.find(x=>x[0]===no);
    if(e) el("kcDetay").innerHTML = `<div class="det" style="margin-top:14px">
      <h2>${e[0]}</h2>${kcTip(e, +ki)}</div>`;
  }
}

function kcTip(e, ki){
  const ekip = (e[13]||[]).map(i=>KYN[i][0]).join(", ");
  return `<b>${e[0]}</b> · ${MAL[e[1]][0]}<br><span class="m">${MAL[e[1]][1]||""}</span><hr>
    <div class="r"><span class="m">Kaynak</span><span>${KYN[ki][0]} ${KYN[ki][1]?"· "+KYN[ki][1].split(" - ")[0]:""}</span></div>
    ${e[3]>=0?`<div class="r"><span class="m">Makine</span><span>${KYN[e[3]][0]}</span></div>`:""}
    ${e[11]?`<div class="r"><span class="m">Kapasite grubu</span><span>${e[11]}</span></div>`:""}
    ${e[12]>1?`<div class="r"><span class="m">Gerekli kişi</span><span>${e[12]} kişi${ekip?" ("+ekip+")":""}</span></div>`:""}
    <div class="r"><span class="m">Tarih</span><span>${e[4]} → ${e[5]} (${e[6]} gün)</span></div>
    <div class="r"><span class="m">Miktar</span><span>${fmt(e[7])}</span></div>
    <div class="r"><span class="m">Yük</span><span>${fmt(e[8],2)} saat</span></div>
    <div class="r"><span class="m">Tetikleyen sipariş</span><span>${e[9]||"–"}</span></div>`;
}


// ================= KAYNAK KAPATMA (izin / ariza / bakim) =================
// Kayitlar kapanislar.json dosyasinda tutulur; veri_aktar.py bu dosyayi okuyup
// plan girdisindeki "Kapanislar" sayfasina yazar, motor da o tarihlerde
// kaynagin kapasitesini kapatir/azaltir.
const KAP_TUR = ["IZIN", "ARIZA", "BAKIM", "DIGER"];
const KAP_TURAD = {IZIN:"İzin", ARIZA:"Arıza", BAKIM:"Bakım", DIGER:"Diğer"};
let kapKaydedildi = true;

function kapGunSay(k){
  const a = d2(k.bas), b = d2(k.bit);
  return Math.max(1, Math.round((b-a)/86400000) + 1);
}

function cizKapatma(){
  const kutu = el("kapatma");
  const bugun = V.ozet["Plan baslangic tarihi"] || HAFTALAR[0];
  const secenek = KYN.map((k,i)=>({i, kod:k[0], ad:k[1]||"", tip:k[2]}))
    .sort((a,b)=> (a.tip||"").localeCompare(b.tip||"") || a.kod.localeCompare(b.kod));

  const satirlar = KAPANIS.map((k,i)=>{
    const ki = KYN.findIndex(x=>x[0]===k.kaynak);
    const tam = k.oran <= 0;
    return `<tr>
      <td><b>${k.kaynak}</b><br><span class="rn">${ki>=0?(KYN[ki][1]||""):"<span class='late'>tanımsız kaynak</span>"}</span></td>
      <td><span class="pill ${tam?"warn":""}">${KAP_TURAD[k.tur]||k.tur}</span></td>
      <td class="nw">${k.bas}</td><td class="nw">${k.bit}</td>
      <td class="n">${kapGunSay(k)}</td>
      <td class="n">${tam ? "<b>kapalı</b>" : "%"+k.oran}</td>
      <td>${k.not||""}</td>
      <td><button data-sil="${i}" title="Sil">✕</button></td></tr>`;
  }).join("");

  kutu.innerHTML = `<h2>Kaynak kapatma — izin, arıza, bakım</h2>
    <p class="note">Bir kişiyi izne çıkarmak ya da bir makine/kalıbı arızaya almak için
      tarih aralığı girin. O aralıkta kaynağın kapasitesi <b>kalan kapasite %</b> kadar
      olur (0 = tamamen kapalı, 50 = yarım gün/yarım kapasite). Kaydettikten sonra planın
      yeniden çalıştırılması gerekir.</p>
    <div class="kapform">
      <label>Kaynak
        <input list="kapListe" id="kapKaynak" placeholder="kod ara… (örn. 1022, KM08)">
        <datalist id="kapListe">${secenek.map(x=>
          `<option value="${x.kod}">${x.kod} — ${x.ad} (${TIPAD[x.tip]||x.tip})</option>`).join("")}</datalist>
      </label>
      <label>Tür
        <select id="kapTur">${KAP_TUR.map(t=>`<option value="${t}">${KAP_TURAD[t]}</option>`).join("")}</select>
      </label>
      <label>Başlangıç<input type="date" id="kapBas" value="${bugun}"></label>
      <label>Bitiş<input type="date" id="kapBit" value="${bugun}"></label>
      <label>Kalan kapasite %<input type="number" id="kapOran" min="0" max="100" value="0"></label>
      <label class="gen">Açıklama<input type="text" id="kapNot" placeholder="yıllık izin, hidrolik arıza…"></label>
      <button class="pri" id="kapEkle">Ekle</button>
    </div>
    ${KAPANIS.length ? `<table class="dt"><thead><tr>
        <th>Kaynak</th><th>Tür</th><th>Başlangıç</th><th>Bitiş</th><th class="n">Gün</th>
        <th class="n">Kalan kapasite</th><th>Açıklama</th><th></th>
      </tr></thead><tbody>${satirlar}</tbody></table>`
      : `<p class="note">Henüz kapatma kaydı yok.</p>`}
    <div class="bar" style="position:static;margin-top:12px">
      <button class="pri" id="kapKaydet">kapanislar.json kaydet</button>
      <button id="kapTemizle">Tümünü sil</button>
      <span class="durum ${kapKaydedildi?"":"err"}" id="kapDurum">${
        KAPANIS.length} kayıt${kapKaydedildi?"":" · kaydedilmedi"}</span>
      <span class="spacer"></span>
      <span class="hint">${V.kapanisKaynagi
        ? "Okunan kaynak: <b>"+V.kapanisKaynagi+"</b>"
        : "Henüz kayıtlı dosya yok"} · Kaydettikten sonra calistir.bat ile planı yenileyin.</span>
    </div>
    <p class="note" style="margin-top:10px">
      <b>Kaydetme yeri önemli:</b> tarayıcı "kaydet" dediğinizde dosyayı çoğu zaman
      <i>İndirilenler</i> klasörüne koyar. Sorun değil — <code>calistir.bat</code>
      her çalıştığında İndirilenler'deki daha yeni <code>kapanislar.json</code>
      dosyasını program klasörüne otomatik taşır. Yine de kaydetme penceresi
      açılırsa program klasörünü seçip mevcut dosyanın üzerine yazmak en temizi.</p>`;

  el("kapEkle").onclick = ()=>{
    const kod = el("kapKaynak").value.trim();
    const bas = el("kapBas").value, bit = el("kapBit").value || el("kapBas").value;
    if(!kod || !bas){ el("kapDurum").textContent = "kaynak ve başlangıç gerekli";
                      el("kapDurum").className = "durum err"; return; }
    if(!KYN.some(x=>x[0]===kod)){ el("kapDurum").textContent =
        `'${kod}' diye bir kaynak yok`; el("kapDurum").className = "durum err"; return; }
    KAPANIS.push({kaynak:kod, bas, bit: bit<bas?bas:bit, tur:el("kapTur").value,
                  oran:Math.max(0,Math.min(100,+el("kapOran").value||0)),
                  not:el("kapNot").value.trim()});
    KAPANIS.sort((a,b)=> a.bas.localeCompare(b.bas) || a.kaynak.localeCompare(b.kaynak));
    kapKaydedildi = false; ciz();
  };
  kutu.querySelectorAll("[data-sil]").forEach(b=>
    b.onclick = ()=>{ KAPANIS.splice(+b.dataset.sil,1); kapKaydedildi=false; ciz(); });
  el("kapTemizle").onclick = ()=>{ KAPANIS.length=0; kapKaydedildi=false; ciz(); };
  el("kapKaydet").onclick = async ()=>{
    await dosyaKaydet("kapanislar.json",
      {kaydedildi:new Date().toLocaleString("tr-TR"), kapanislar:KAPANIS}, el("kapDurum"));
    kapKaydedildi = true;
  };
}

// ---- sekmeler
function sekmeSec(ad){
  document.querySelectorAll("#anaTab button").forEach(b=>
    b.setAttribute("aria-selected", String(b.dataset.t===ad)));
  document.querySelectorAll("section.sek").forEach(x=>
    x.hidden = x.dataset.t !== ad);
  ciz();                     // gorunur hale gelen bolum yeniden olculerek cizilsin
}
document.querySelectorAll("#anaTab button").forEach(b=>
  b.onclick = ()=> sekmeSec(b.dataset.t));

// ---- plan baslangicinin haftasindan ac
const pb = V.ozet["Plan baslangic tarihi"];
if(pb){ const t = iso(pzt(d2(pb))); const i = HAFTALAR.indexOf(t); if(i>=0) hi=i; }
selH.value = hi;
ciz();
addEventListener("keydown", e=>{
  if(e.target.tagName==="INPUT"||e.target.tagName==="SELECT") return;
  if(e.key==="ArrowLeft") el("prev").click();
  if(e.key==="ArrowRight") el("next").click();
});
</script>
</body>
</html>
"""


if __name__ == "__main__":
    g = sys.argv[1] if len(sys.argv) > 1 else "MRP_Sonuc.xlsx"
    c = sys.argv[2] if len(sys.argv) > 2 else "MRP_Cizelge.html"
    p = sys.argv[3] if len(sys.argv) > 3 else "MRP_Plan_Girdi.xlsx"
    yol, nk, nkyn, nmal = olustur(g, c, p)
    print("Cizelge yazildi : %s" % yol)
    print("   %d uretim kaydi | %d kaynak | %d urun" % (nk, nkyn, nmal))

# -*- coding: utf-8 -*-
"""
AKIS (flow) verisi: haftalik malzeme akisini urun agaci yapisina gore cikarir.

Cizelge HTML'indeki "Siparis akisi" bolumu bu veriyi kullanir. Her hafta icin:
  n : dugumler  [mal_ix, acilis, uretim, tuketim, sevk, kapanis, ie_adet, ie_saat,
                 giren]   (giren = stok defterine o hafta giren; uretim ertesi
                 gun stoga gectigi icin hafta sonunda farkli olabilir)
  e : kenarlar  [ust_ix, alt_ix, miktar]      (ust 1 adet icin alt 'oran' adet)
  s : siparisler[belge, kalem, mal_ix, miktar, teslim, durum, bitis, gecikme]

Stok sayilari MRPSonuc stok defterinden birebir okunur (uydurma yok):
acilis = haftanin ilk hareketinin kumulatif stogu eksi o hareketin miktari.
"""

import datetime as dt
from collections import defaultdict


def _g(v):
    return v.date() if isinstance(v, dt.datetime) else v


def _pzt(g):
    return g - dt.timedelta(days=g.weekday())


def urun_agaci(pg):
    """
    UrunAgaciSeviyeli (girinti seviyeli liste) -> kenarlar ve seviyeler.
    Dondurur: (kenarlar {(ust, alt): oran}, seviye {malzeme: en derin seviye})
    """
    kenar, seviye = {}, {}
    if pg is None or "UrunAgaciSeviyeli" not in pg.sheetnames:
        return kenar, seviye
    ws = pg["UrunAgaciSeviyeli"]
    # kolonlari basliktan bul (sayfaya aciklama gibi kolonlar eklenebilir)
    j = None
    satirlar = ws.iter_rows(values_only=True)
    for r in satirlar:
        if r and "Seviye" in r and "Malzeme" in r:
            j = {b: i for i, b in enumerate(r) if b}
            break
    if j is None:
        return kenar, seviye
    i_s, i_m = j["Seviye"], j["Malzeme"]
    i_o = j.get("Miktar (alt/1 ust)", j.get("Miktar"))
    yigin = {}                                   # seviye -> o seviyedeki son malzeme
    for r in satirlar:
        if r[i_s] is None or not r[i_m]:
            continue
        try:
            sv = int(r[i_s])
        except (TypeError, ValueError):
            continue
        kod = str(r[i_m]).strip()
        oran = float((r[i_o] if i_o is not None else 1) or 1)
        yigin[sv] = kod
        seviye[kod] = max(seviye.get(kod, 0), sv)
        if sv > 0 and (sv - 1) in yigin:
            ust = yigin[sv - 1]
            if ust != kod:
                kenar[(ust, kod)] = oran
    return kenar, seviye


def stok_defteri(wb):
    """
    MRPSonuc -> malzeme -> hafta(pzt) -> [acilis, giren, tuketim, sevk, kapanis]

    "Baslangic Stogu" satiri bir HAREKET degil, ilk haftanin ACILISIDIR; hareket
    sayilirsa ilk hafta acilisi sifir gorunur.
    """
    if "MRPSonuc" not in wb.sheetnames:
        return {}
    ws = wb["MRPSonuc"]
    it = ws.iter_rows(values_only=True)
    bas = list(next(it))
    j = {b: i for i, b in enumerate(bas)}
    i_m, i_b = j.get("Malzeme", 0), j.get("Belge", 2)
    i_q, i_k = j.get("Miktar", 3), j.get("Kumulatif Stok", 4)
    i_t = j.get("Tarih", 5)

    d = defaultdict(dict)
    for r in it:
        kod, belge, mik, kum, tar = r[i_m], r[i_b], r[i_q], r[i_k], r[i_t]
        if not kod or tar is None or mik is None or kum is None:
            continue
        b = str(belge or "")
        if b.startswith("Aciklama"):
            continue
        h = _pzt(_g(tar)).isoformat()
        kutu = d[kod].get(h)
        if kutu is None:
            # acilis = ilk hareketin oncesindeki stok; "Baslangic Stogu"
            # satirinda ise stogun kendisi acilistir
            ac = float(kum) if b.startswith("Baslangic") else float(kum) - float(mik)
            kutu = [ac, 0.0, 0.0, 0.0, ac]
            d[kod][h] = kutu
        if b.startswith("Uretim") or b.startswith("Acik"):
            kutu[1] += float(mik)
        elif b.startswith("Ikincil"):
            kutu[2] += -float(mik)
        elif b.startswith("Satis"):
            kutu[3] += -float(mik)
        kutu[4] = float(kum)
    return d


def haftalik_uretim(wb):
    """
    UretimPlani -> malzeme -> hafta(pzt) -> o hafta URETILEN miktar.
    Stok defterindeki giris, aktarma kurali geregi uretimden bir gun sonradir;
    kutularda is emirleriyle ayni sayiyi gostermek icin uretim gunu kullanilir.
    """
    u = defaultdict(lambda: defaultdict(float))
    if "UretimPlani" not in wb.sheetnames:
        return u
    ws = wb["UretimPlani"]
    it = ws.iter_rows(values_only=True)
    j = {b: i for i, b in enumerate(next(it))}
    i_t, i_m, i_q = j.get("Tarih", 0), j.get("Malzeme", 1), j.get("Miktar", 9)
    for r in it:
        g = _g(r[i_t])
        if not isinstance(g, dt.date) or not r[i_m]:
            continue
        u[r[i_m]][_pzt(g).isoformat()] += float(r[i_q] or 0)
    return u


def akis_verisi(wb, pg, mlz, emirler, mal_ad):
    """
    mlz      : OrderedDict malzeme -> aciklama (cizelgenin malzeme listesi; buyur)
    emirler  : cizelgenin is emri listesi ([0]=no, [1]=mal_ix, [4]=bas, [7]=adet, [8]=saat)
    mal_ad   : malzeme -> aciklama (plan girdisinden, eksikleri tamamlamak icin)
    Dondurur : {"hafta": {...}, "seviye": {mal_ix: seviye}, "tip": {mal_ix: tip}}
    """
    kenar, seviye = urun_agaci(pg)
    defter = stok_defteri(wb)
    uretim = haftalik_uretim(wb)

    # --- siparisler
    sip = defaultdict(list)
    if "SiparisOzeti" in wb.sheetnames:
        ws = wb["SiparisOzeti"]
        it = ws.iter_rows(values_only=True)
        j = {b: i for i, b in enumerate(next(it))}
        for r in it:
            if not r[j["Belge"]]:
                continue
            t = _g(r[j["Teslim Tarihi"]])
            if not isinstance(t, dt.date):
                continue
            bit = _g(r[j["Ileri Plan: Bitis"]])
            sip[_pzt(t).isoformat()].append({
                "belge": str(r[j["Belge"]]), "kalem": r[j["Kalem"]],
                "mal": r[j["Malzeme"]], "miktar": float(r[j["Miktar"]] or 0),
                "teslim": t.isoformat(),
                "durum": str(r[j["Durum"]] or ""),
                "bitis": bit.isoformat() if isinstance(bit, dt.date) else "",
                "gecikme": int(r[j["Gecikme (gun)"]] or 0)})

    # --- is emri sayisi / saati (malzeme x hafta)
    ie = defaultdict(lambda: [0, 0.0])
    ters = list(mlz)
    mix = {k: i for i, k in enumerate(ters)}
    for e in emirler:
        h = _pzt(dt.date.fromisoformat(e[4])).isoformat()
        kod = ters[e[1]] if e[1] < len(ters) else None
        if kod:
            ie[(kod, h)][0] += 1
            ie[(kod, h)][1] += e[8]

    def ix(kod):
        if kod not in mix:
            mlz[kod] = mal_ad.get(kod, "")
            mix[kod] = len(mix)
        return mix[kod]

    haftalar = sorted({h for v in defter.values() for h in v} | set(sip))
    cikti = {}
    for h in haftalar:
        aktif = {k for k, v in defter.items()
                 if h in v and any(abs(x) > 1e-6 for x in v[h][1:4])}
        aktif |= {k for k, v in uretim.items() if v.get(h, 0) > 1e-6}
        aktif |= {s["mal"] for s in sip.get(h, [])}
        if not aktif:
            continue
        # dugumler
        dugum, kodlar = [], sorted(aktif)
        for kod in kodlar:
            a, gir, t, sv, kap = defter.get(kod, {}).get(h, [0.0] * 5)
            n = ie.get((kod, h), [0, 0.0])
            ure = uretim.get(kod, {}).get(h, 0.0)
            dugum.append([ix(kod), round(a, 2), round(ure, 2), round(t, 2),
                          round(sv, 2), round(kap, 2), n[0], round(n[1], 1),
                          round(gir, 2)])
        kset = set(kodlar)
        kenarlar = []
        for (ust, alt), oran in kenar.items():
            if ust in kset and alt in kset:
                uret = uretim.get(ust, {}).get(h, 0.0)
                kenarlar.append([ix(ust), ix(alt), round(uret * oran, 2)])
        siparisler = [[s["belge"], s["kalem"], ix(s["mal"]), round(s["miktar"], 2),
                       s["teslim"], s["durum"], s["bitis"], s["gecikme"]]
                      for s in sip.get(h, [])]
        cikti[h] = {"n": dugum, "e": kenarlar, "s": siparisler}

    sev = {}
    for kod, i in mix.items():
        if kod in seviye:
            sev[str(i)] = seviye[kod]
    return {"hafta": cikti, "seviye": sev}

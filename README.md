# MRP ve Paralel Planlama Programı

Debak'ın ERP verisinden (ürün ağacı, stok, sipariş, personel, kalıp–makine) **malzeme ihtiyaç planı** ve **sonlu kapasiteli üretim planı** çıkaran program. ERP'nin MRP'sinden farkı: alt seviye bitmeden üst seviye başlamıyor diye beklemez; **seviyeler paralel ilerler** (bir seviyenin bugünkü üretimi yarın üst seviyede kullanılabilir). Sonuç haftalık iş emirleri, kaynak doluluğu ve sipariş teslim durumu olarak Excel'e ve tek sayfalık bir HTML çizelgeye yazılır.

Programın nasıl geliştiğini, hangi kararın neden alındığını görmek için: **[SOHBET_GECMISI.md](SOHBET_GECMISI.md)**.

---

## Çalıştırma

Windows'ta program klasöründe **`calistir.bat`**'a çift tıklayın. Gerekenler: Python 3.10+ ve `openpyxl` (bat dosyası yoksa kurar).

`calistir.bat` dört adımı sırayla çalıştırır:

| Adım | Script | Ne yapar |
|---|---|---|
| 0 | `ayar_topla.py` | Çizelgeden kaydedilip İndirilenler'e düşen `ayarlar.json` / `atamalar.json` / `kapanislar.json` dosyalarının daha yenisini program klasörüne taşır |
| 1 | `veri_aktar.py` | Ham ERP sayfalarını (`MRP_Girdi.xlsx`) plan girdisine çevirir → `MRP_Plan_Girdi.xlsx` |
| 2 | `mrp_planlama.py` | MRP patlatma + geri planlama + günlük ileri simülasyon + 9 kontrol → `MRP_Sonuc.xlsx` |
| 3 | `cizelge_olustur.py` (+ `akis_veri.py`) | Haftalık görsel çizelge → `MRP_Cizelge.html` |

Elle çalıştırmak isterseniz:

```bat
python ayar_topla.py
python veri_aktar.py MRP_Girdi.xlsm MRP_Plan_Girdi.xlsx
python mrp_planlama.py MRP_Plan_Girdi.xlsx MRP_Sonuc.xlsx
python cizelge_olustur.py MRP_Sonuc.xlsx MRP_Cizelge.html MRP_Plan_Girdi.xlsx
```

Plan başlangıç tarihi `ayarlar.json › plan.baslangic` alanından gelir (komut satırındaki tarihi ezer). **Stok dökümü hangi gün alındıysa başlangıç o gün olmalı**; aksi hâlde aradaki günlerde yapılmış üretim/sevkiyat görünmez.

---

## Dosyalar

### Girdi — `MRP_Girdi.xlsx` (ERP'den, elle güncellenir)

| Sekme | İçerik |
|---|---|
| `URUN AGACI REVIZYONLU` | Seviyeli ürün ağacı: Seviye · **Üst Malzeme** · Malzeme Kodu · Açıklama · Miktar · Birim · İş Merkezi / Kapasite Grubu · Çalışan Personel Sayısı · Rota Süresi (sn) · **FIRINALAMA**. Yoksa eski `URUN AGACI SEVIYELI` okunur. |
| `PERSONELLISTESI` | Görünen Ad · Sicil No · Görev · KAPASITE GRUPLARI 1–4 (bir kişi birden fazla grupta olabilir) |
| `PAKETLEME` | Paketleme grubu → paketçi kodları |
| `STOK` | CANIAS stok dökümü (yalnız 100/200/300 depoları); `Kullanılabilir Stok` ve `Depo` kullanılır. 1. satırdaki **CANIAS'tan Güncelle** butonuyla yenilenir |
| `SİPARİŞ` | Açık satış siparişleri (CANIAS SALT13 raporu); `Açık Miktar` ve `Teslim Tarihi` kullanılır. 1. satırdaki **CANIAS'tan Güncelle** butonuyla yenilenir |

Dosya artık makrolu: **`MRP_Girdi.xlsm`** (`calistir.bat` önce bunu, yoksa `MRP_Girdi.xlsx`'i kullanır). STOK ve SİPARİŞ'te 1. satır buton satırıdır, başlıklar 2. satırdadır.

**CANIAS güncelleme butonu:** Butona basınca gizli `CANIAS` sayfasındaki sorgu çalışır, sekme baştan yazılır ve dosya kaydedilir. İlk basışta SQL Server giriş penceresi açılır (CANIAS rapor kullanıcısının şifresi); şifre kaydedilmez, Excel kapanana kadar hatırlanır.

- Sorgular: `stok_sorgu.sql`, `siparis_sorgu.sql`; bağlantı: `canias_baglanti.json`. Değiştirdikten sonra Excel kapalıyken `python excel_kurulum.py` çalıştırın (sorguları gizli sayfaya yükler, butonlar zaten varsa dokunmaz, önce `yedek/`'e kopya alır).
- Makro kodu `canias_guncelle.bas`. Excel güvenlik ayarı nedeniyle **bir kez elle** içe aktarılır: dosyayı açın → Alt+F11 → File › Import File… → `canias_guncelle.bas` → kapatıp kaydedin. Açarken "İçeriği Etkinleştir" demek gerekir.
- Dosyayı openpyxl ile **yazan** bir script çalıştırmayın: butonlar ve makro silinir (okumak sorun değil).

### Ayar dosyaları (çizelgenin içinden düzenlenir)

| Dosya | İçerik |
|---|---|
| `ayarlar.json` | Plan parametreleri, varsayılan çalışma saatleri, operatör katsayısı, kaynak bazlı istisnalar |
| `kapanislar.json` | Kaynak kapatmaları: izin, arıza, bakım (tarih aralığı + kalan kapasite %) |
| `atamalar.json` | Planlamacının iş emirlerine elle seçtiği gerçek kişiler (plana etki etmez, kayıt tutar) |

### Çıktılar

| Dosya | İçerik |
|---|---|
| `MRP_Plan_Girdi.xlsx` | Motorun okuduğu düzenli girdi: Malzemeler, UrunAgaciSeviyeli, Kaynaklar, Rotalar, KalipMakine, Atamalar, Talepler, Kapanislar, Gruplar, Parametreler, **AktarimRaporu** (veri uyarıları burada) |
| `MRP_Sonuc.xlsx` | Ozet, SiparisOzeti, UretimPlani, **IsEmirleri**, KalipMakineAtama, GeriPlan, KapasiteYuku, MRPSonuc (stok defteri), HammaddePlani, **Kontrol** |
| `MRP_Cizelge.html` | Tarayıcıda açılan çizelge — sekmeler: Kaynak doluluğu · Kaynak çizelgesi (Gantt) · İş emirleri · Sipariş akışı (ürün ağacı üzerinde) · Gecikmeler · Kaynak kapatma · ⚙ Ayarlar |

---

## Planlama kuralları (güncel)

**Malzeme ve ağaç**

- Ürün ağacında ebeveyn **"Üst Malzeme" kolonundan** alınır (satır sırası karışık olsa bile doğru bağlanır). Üst Malzeme sayfada olmayan bir kodsa girintiye bakılır; seviye atlamış ve Üst Malzeme de geçersizse satır atlanır. Hepsi `AktarimRaporu`'na yazılır.
- Aynı bağ farklı miktarla iki kez geçerse ilki kullanılır ve uyarı verilir.
- Stok: `Kullanılabilir Stok`; **400 ve 500 depoları plana girmez** (`ayarlar.json › plan.haric_depolar`).

**Kaynaklar**

- Bir malzemenin birden fazla kaynağı olabilir; **1. öncelik dolunca kalan aynı gün 2.'ye taşar**.
- Enjeksiyon adımı bir **kalıptır**; kalıp yalnızca onaylı makinelerinde, öncelik sırasıyla çalışır ve kalıp + makine aynı anda kapasite tüketir. Bir kalıp aynı anda tek makinede olabilir.
- Enjeksiyon operatörleri **ortak havuz** (1 operatör = 1 makine, günde 8 saat, cumartesi dahil). Başka bir kapasite grubunda bireysel planlanan operatör havuza sayılmaz.
- Kişiler günde 9 saat, hafta içi; makine ve kalıplar 24 saat, cumartesi dahil.
- **Çalışan personel sayısı** = işi aynı anda yapan kişi sayısı; o kadar kişi aynı gün aynı işe ekip olarak ayrılır.
- **Fırınlar partili çalışır**: `FIRINALAMA` = `20000 AD/ 20 H` → bir fırın 20 saatte en fazla 20.000 adet. Yarım parti de tam süre tüketir; aynı gün aynı fırına gelen aynı ürün açık partiyi doldurur; yarım parti ancak işin kalanının tamamı hazırsa ya da teslim yaklaştıysa başlatılır.
- **Kaynak kapatma**: `kapanislar.json` (çizelgedeki *Kaynak kapatma* sekmesi) — o tarihlerde kaynağın kapasitesi `kalan kapasite %` olur.

**Dağıtım**

- **Günlük iş çeşidi = 1**: bir kişi gün içinde tek ürüne çalışır.
- **İş başına azami kaynak = 1**: bir ürün bir haftada tek kişiye/kalıba verilir, erken bitsin diye bölünmez.
- **Takviye**: iş tek kişiyle teslime yetişmeyecekse gerektiği kadar kişi eklenir (2 iş günü emniyet payı).
- Teslim tarihi gelmiş işlerde bu sınırlar kalkar.
- **Erken sevkiyat yok**: mal erken biterse stokta bekler, teslim gününde sevk edilir.

Hepsi çizelgede **⚙ Ayarlar** penceresinden değiştirilebilir; kaydedince `ayarlar.json` yazılır, `calistir.bat` ile plan yenilenir.

**Kontroller** (`MRP_Sonuc.xlsx › Kontrol`): kapasite aşımı yok · tatilde üretim yok · üretim malzemesinde negatif stok yok · planlı emirlerin tamamı üretildi · sevkiyat = talep · seviyeler paralel · aktarma kuralı korundu · alternatif önceliğine uyuldu · kalıplar sadece tanımlı makinelerde.

---

## Bilinen veri eksikleri (25.11.2026, 16:49 girdisiyle)

`MRP_Plan_Girdi.xlsx › AktarimRaporu` her çalıştırmada güncel listeyi verir. Son durumda:

- Siparişi olup ürün ağacı olmayan mamuller: MAM005-00086, MAM005-00087, MAM02029, MAM100 — toplam 32.901 adet plana girmiyor.
- İş merkezi boş lebleme adımları: LEB069-00023, LEB069-00024 — kapasite sınırı olmadan planlanıyor (`KAYNAK-TANIMSIZ`; AktarimRaporu'nda uyarı olarak çıkmıyor, İş Emirleri'nde görünür).
- `CAPROBOTU` ve `BOYROBOTU` personel grubu olarak tanımlı (1078 ve 1032); robotlar bu kişilerin takvimiyle (günde 9 saat, hafta içi) çalışıyor. Robotlar kesintisiz çalışıyorsa makine olarak tanımlanmaları daha doğru olur.
- Bilgi: `HAM001-00009-R00` (ENJ003-00025 altı) sayfada yanlış yerde duruyor; Üst Malzeme kolonu doğru olduğu için bağ doğru kuruluyor.

25.11.2026'da giderilenler: ENJ001-00086 boş satırı, ENJ001-00088 personel sayısı, FRN003-00046 / FRN069-00022 / ENJ069-00022 / ENJ009-00007 iş merkezleri, rota süresi 0 olan adımlar, FRN069-00021 fırın partisi, ATT003-00029 → ATT003-00058 Üst Malzeme düzeltmesi.

---

## Araçlar

- `excel_kurulum.py` — `MRP_Girdi`'ye CANIAS butonlarını kurar ve sorguları gizli `CANIAS` sayfasına yükler (yukarıya bakın). **Not:** stok sorgusu yalnız 100/200/300 depolarını çeker; 101, 401, 907, 950 stokları plana girmez.
- `SOHBET_GECMISI.md` — çalışma oturumlarının özet kaydı; şimdilik elle güncelleniyor (otomatik üretici henüz yok).
- `girdi_olustur.py` — boş/örnek girdi şablonu üretir. **Dikkat:** dosya adı vermezseniz `MRP_Girdi.xlsx`'in üzerine yazar; her zaman farklı bir ad verin: `python girdi_olustur.py ornek.xlsx --ornek`.

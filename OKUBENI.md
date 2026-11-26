# MRP ve Paralel Planlama Programı

Çok ürünlü, çok seviyeli MRP patlatma + **sonlu kapasiteli, seviyeleri paralel çalıştıran** üretim planlama.

## Kurulum

1. Python 3 kurulu olmalı (kurulumda **"Add Python to PATH"** işaretli olsun).
2. Dosyaları bir klasöre koy.
3. `calistir.bat` dosyasına çift tıkla. Eksik kütüphane varsa kendisi kurar.

Komut satırından:

```
python mrp_planlama.py MRP_Girdi.xlsx MRP_Sonuc.xlsx
python girdi_olustur.py YeniDosya.xlsx --bos      # boş şablon
python girdi_olustur.py Ornek.xlsx     --ornek    # çok ürünlü örnek veri
```

---

## Neden "paralel"?

ERP çıktısında her seviye, alt seviye **tamamen bittikten sonra** başlıyordu.
Bu programda seviyeler **eşzamanlı** akar:

> Alt seviye bir günlük üretimini tamamladığında, o miktar ertesi gün üst seviyenin
> kullanımına açılır. Üst seviye alt seviyenin bitmesini beklemez.

`Parametreler` > **Aktarma Kuralı** ile değiştirilebilir:

| Değer | Anlamı |
|---|---|
| `GUNLUK` | Alt seviyenin 1 günlük üretimi birikince üst seviye başlar *(varsayılan)* |
| `ADET` | Adet adet akış — aynı gün içinde de aktarım olur, en kısa süre |
| `SABIT` | `Aktarma Partisi` kadar (örn. bir kasa) birikince |

---

## Bir malzeme için birden fazla kaynak

`Rotalar` sayfasında bir malzemeye öncelik sırasıyla birden fazla kaynak tanımlarsın; **her alternatifin kendi saniye/adet'i** olur:

| Malzeme | Öncelik | Kaynak | Saniye/Adet |
|---|---|---|---|
| URN-100 | 1 | KISI-1 | 12 |
| URN-100 | 2 | KISI-2 | 14 |
| ENJ-500 | 1 | KALIP-500A | 136 |
| ENJ-500 | 2 | KALIP-500B | 150 |

**Taşma kuralı:** 1. öncelikli kaynağın o günkü kapasitesi dolduğunda kalan miktar **aynı gün** 2. önceliğe geçer. Örnek: KISI-1 günde 9 saat × 3600 ÷ 12 sn = 2.700 adet yapar; 4.000 adetlik ihtiyaçta kalan 1.300 adet KISI-2'ye taşar ve orada 14 sn'den 5,06 saat sürer.

Bu sayfada yer almayan malzemeler `Malzemeler` sayfasındaki tek kaynağı kullanmaya devam eder — mevcut dosyaların bozulmaz.

---

## Kalıp ve makine

Enjeksiyon adımı **kalıbı** temsil eder. Bir parçanın birden fazla kalıbı olabilir (`Rotalar`'da öncelik sırasıyla), her kalıp birden fazla makinede çalışabilir (`KalipMakine`'de öncelik sırasıyla):

| Kalıp | Öncelik | Makine | Hız Çarpanı |
|---|---|---|---|
| KALIP-500A | 1 | PRES-1 | 1,00 |
| KALIP-500A | 2 | PRES-2 | 1,10 |
| KALIP-500B | 1 | PRES-2 | 1,00 |

- Üretim **hem kalıbın hem makinenin** kapasitesini aynı anda tüketir.
- **Makineler tüm ürünler için ortaktır** — farklı ürünlerin kalıpları aynı presi paylaşır, kapasite ortak havuzdan gider.
- **Hız Çarpanı**: 1,10 = o makinede %10 yavaş (süre = Saniye/Adet × Çarpan).

### Kalıp yapışkan bağlanır

Kalıp bir makineye bağlandığında **orada kalır** — gereksiz yere taşınmaz. Sadece o makine ona yer vermediğinde sıradaki makineye geçer, ve orada da kalmaya devam eder.

`Parametreler` > **Minimum Kalıp Kalış (gün)** ile bunu zorlayabilirsin:

| Değer | Davranış |
|---|---|
| `0` | Serbest — yer bulamazsa hemen taşınır *(varsayılan)* |
| `7` | Bağlandığı makinede en az bir hafta kalır; o makine kapalıysa/doluysa bekler, fazla stok yapar ama taşınmaz |

Çıktıdaki **KalipMakineAtama** sayfası hangi kalıbın hangi tarihte hangi makinede olduğunu, kaç gün çalıştığını ve kaç kez taşındığını gösterir.

---

## Çok ürünlü çalışma

- **Malzemeler** sayfasında nihai ürün, yarı mamul ve hammadde aynı listede — kaç tane olursa olsun.
- **Talepler** sayfasında farklı malzemelere ayrı ayrı sipariş girilebilir; aynı malzemenin birden fazla teslim kalemi olabilir.
- Ürün ağacı **dallanabilir** (bir ürünün birden fazla alt parçası) ve bir parça **birden fazla üründe** kullanılabilir. Örneğin aynı enjeksiyon gövdesini iki farklı mamul kullanabilir — ihtiyaçlar toplanır, stok ve kapasite paylaşılır.
- **Birden fazla malzeme aynı kaynağı gösterebilir** — o tezgâhın günlük kapasitesi aralarında paylaşılır, işler teslim tarihine göre (EDD) sıraya girer.

`MRP_ORNEK_CokUrunlu.xlsx` bunların hepsini gösteren çalışan bir örnek: 3 nihai ürün; iki mamulün paylaştığı ENJ-500 gövdesi; ENJ-500 için iki kalıp; iki presi paylaşan üç kalıp; montajda 1. ve 2. öncelikli kişi; açık iş emri ve negatif rezervasyon.

---

## Ürün ağacını iki formatta girebilirsin

**1. Üst-alt çifti** (`UrunAgaci` sayfası) — her satır bir bağlantı:

| Üst Malzeme | Alt Malzeme | Miktar |
|---|---|---|
| URN-300 | YRM-310 | 1 |
| URN-300 | YRM-900 | 3 |
| YRM-310 | HAM-001 | 0,06 |

**2. Seviyeli girintili liste** (`UrunAgaciSeviyeli` sayfası) — ERP ağaç görünümüne benzer:

| Seviye | Malzeme | Miktar |
|---|---|---|
| 0 | URN-300 | |
| 1 | YRM-310 | 1 |
| 2 | HAM-001 | 0,06 |
| 1 | YRM-900 | 3 |

İkisi aynı ağacı tanımlar. Hangisi kolayına geliyorsa onu kullan; ikisini aynı dosyada karışık da kullanabilirsin (bir ürünü çift, diğerini seviyeli). Aynı bağlantı iki formatta da yazılırsa program tekrar saymaz; **oranlar çelişirse hata verir.**

---

## ERP'den kopyala-yapıştır

Program başlık satırını kendi bulur — kolon sırası ve fazladan kolonlar önemli değil.
Bilinen başlıklar otomatik tanınır:

`Malzeme` · `Stok Kodu` · `Malzeme Açıklaması` · `Sal. Bel. Num.` · `KalemNo` ·
`Teslim Tarihi` · `Termin` · `Miktar Br.` · `adet/saniye` · `ANLIK STOK` · `İş Merkezi` · ...

Tanınmayan bir başlık varsa **KolonEslestirme** sayfasına bir satır ekle:

| Sayfa | Dosyadaki Başlık | Program Alanı |
|---|---|---|
| Talepler | Sevk Tar. | Teslim Tarihi |
| Talepler | Parça | Malzeme Kodu |

Bir sayfada dolu satır var ama hiçbiri okunamıyorsa program **sessizce boş plan üretmez** — hangi başlıkları gördüğünü yazıp durur.

---

## Girdi dosyası sayfaları

| Sayfa | İçerik |
|---|---|
| **Kullanim** | Kısa kullanım notları |
| **Malzemeler** | Kod, açıklama, tip (URETIM/SATINALMA), kaynak, saniye/adet, anlık stok, emniyet stoğu |
| **UrunAgaci** | Üst → alt bağlantısı, oran, fire % |
| **UrunAgaciSeviyeli** | Aynı bilginin girintili hali |
| **Kaynaklar** | Kişi / kalıp / makine — hepsi aynı listede, `Tip` kolonuyla ayrılır |
| **Rotalar** | Malzeme → öncelikli kaynak alternatifleri, her biri kendi saniye/adet'i ile |
| **KalipMakine** | Kalıp → öncelikli makineler, hız çarpanı |
| **Talepler** | Satış siparişleri: belge, kalem, malzeme, miktar, teslim tarihi, öncelik |
| **AcikSiparisler** | Açık iş emri / satın alma (+), rezervasyon (−) |
| **Parametreler** | Plan başlangıcı, aktarma kuralı, ufuk, sıralama, tatiller |
| **KolonEslestirme** | ERP başlığı → program alanı |

Sarı hücreler veri alanı; istediğin kadar satır ekleyebilirsin.

---

## Çıktı dosyası sayfaları

| Sayfa | İçerik |
|---|---|
| **Ozet** | Genel durum, kaynak doluluk oranları, **ürün bazında özet**, uyarılar |
| **SiparisOzeti** | Kalem bazında: geri planda gereken başlangıç, ileri planda bitiş, gecikme |
| **UretimPlani** | İleri plan — gün × malzeme × **kaynak × makine** üretim miktarı, yük, doluluk % |
| **KalipMakineAtama** | Hangi kalıp hangi makinede, hangi tarihler arası, kaç kez taşındı |
| **GeriPlan** | Teslim tarihinden geriye plan; geçmişe düşenler işaretli |
| **KapasiteYuku** | Kaynak × gün kapasite / yük / boş saat |
| **MRPSonuc** | ERP formatında hareket listesi + kümülatif stok |
| **HammaddePlani** | Satın alınan malzemelerin giriş/çıkış ve kalan stok akışı |
| **SenaryoAnalizi** | Darboğaza kademeli tezgâh eklenirse ne olur |
| **Kontrol** | Otomatik doğrulama testleri |

---

## Planlama nasıl çalışıyor?

**1. MRP patlatma** — Ürün ağacı çok seviyeli patlatılır. Her seviyede anlık stok, emniyet stoğu ve açık siparişler netlenir; kalan için planlı emir açılır.

**2. Geri planlama** — Her sipariş teslim tarihinden geriye, sonlu kapasiteyle yüklenir. Üst seviye hangi gün ne kadar üretiyorsa alt seviye o miktarı bir gün öncesine kadar hazırlamalıdır. Başlangıç bugünden önceye düşerse `GECMISTE` işaretlenir.

**3. İleri planlama** — Plan başlangıcından itibaren gün gün simülasyon. Her kaynak o günkü kapasitesi kadar üretir, malzeme yoksa bekler. Gerçekçi bitiş tarihleri ve gecikmeler buradan çıkar.

---

## Doğrulama

Her çalıştırmada 9 otomatik kontrol yapılır ve `Kontrol` sayfasına yazılır:

1. Hiçbir günde kapasite aşılmadı *(paylaşımlı kaynaklar dahil)*
2. Tatil / çalışılmayan günde üretim yok
3. Üretim malzemelerinde negatif stok yok
4. Planlı emirlerin tamamı üretildi
5. Sevkiyat = talep
6. Seviyeler paralel çalışıyor (üst üste biniyor)
7. Aktarma kuralı korundu
8. Alternatif kaynak önceliğine uyuldu *(düşük öncelikli kaynak ancak üstü dolduğunda devreye girdi)*
9. Kalıplar yalnızca tanımlı makinelerinde çalıştı

Bir kontrol `HATA` verirse çıktıya güvenme — girdi verisini kontrol et.

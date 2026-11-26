-- CANIAS acik satis siparisleri -> MRP_Girdi.xlsm > SİPARİŞ sekmesi
-- Sekmedeki "CANIAS'tan Guncelle" butonu bu sorguyu calistirir. Kolon adlari (AS [..])
-- dogrudan sekme basliklari olur; veri_aktar.py "Malzeme", "Açık Miktar", "Teslim Tarihi",
-- "Sal. Bel. Num.", "KalemNo", "Harici Sip. No", "Miktar Br." basliklarini okur.
-- Degistirdikten sonra:  python excel_kurulum.py   (sorguyu calisma kitabina yukler)
--
-- Kaynak: CANIAS SALT13 (satis siparis raporu) trace'i, 25.11.2026. Rapordaki kriterler aynen
-- korundu: acik/kismi kalemler (ORDSTAT < 2), silinmemis, durdurulmamis, tesis 01.
-- Rapordaki musteri ara toplam satirlari (ORDSTAT 99) burada uretilmez.
-- Kalem Toplami belgenin kendi para biriminde (P. Br.) gelir; rapor bunu EUR'ya ceviriyordu.
SELECT B.PRODDATE                              AS [Teslim Tarihi],
       A.CUSTOMER                              AS [Müşteri],
       A.NAME1                                 AS [Müşteri Adı],
       B.MATERIAL                              AS [Malzeme],
       B.MTEXT                                 AS [Malzeme Adı],
       B.QUANTITY                              AS [Miktar],
       B.QUNIT                                 AS [Miktar Br.],
       B.AVAILQTY * B.FACTOR                   AS [Açık Miktar],
       B.SKUNIT                                AS [SBr],
       B.SUBTOTAL                              AS [Kalem Toplamı],
       A.CURRENCY                              AS [P. Br.],
       CONCAT(A.PRINTEDORDTYPE, '-', A.PRINTEDORDNUM) AS [Harici Sip. No],
       A.VALIDFROM                             AS [Belge Tarihi],
       CASE B.ORDSTAT WHEN 0 THEN N'Açık' WHEN 1 THEN N'Kısmi' ELSE CAST(B.ORDSTAT AS varchar(5)) END
                                               AS [Bel. Statüsü],
       A.DOCTYPE                               AS [Satış Belge Tipi],
       A.DOCNUM                                AS [Sal. Bel. Num.],
       B.ITEMNUM                               AS [KalemNo],
       B.VOPTIONS                              AS [Seçenek açıklaması],
       B.PLANT                                 AS [Tesis],
       C.MATTYPE                               AS [Malzeme Tipi],
       C.MATGRP                                AS [Mlz.Grb.]
  FROM IASSALHEAD A, IASSALITEM B, IASMATBASIC C, IASBAS030, IASBAS041
 WHERE A.CLIENT = '00'
   AND A.COMPANY = '01'
   AND A.ISORDCHAR = 1
   AND A.ISDELETE = 0
   AND B.ORDSTAT < 2
   AND B.CLIENT = A.CLIENT
   AND B.COMPANY = A.COMPANY
   AND B.DOCTYPE = A.DOCTYPE
   AND B.DOCNUM = A.DOCNUM
   AND B.SETITEMNUM = 0
   AND B.BUSAREA = '*'
   AND B.PLANT = '01'
   AND B.ISSTOP = 0
   AND C.CLIENT = B.CLIENT
   AND C.COMPANY = B.COMPANY
   AND C.MATERIAL = B.MATERIAL
   AND C.VALIDFROM <= CAST(GETDATE() AS date)
   AND C.VALIDUNTIL >= CAST(GETDATE() AS date)
   AND IASBAS030.CLIENT = C.CLIENT
   AND IASBAS030.COMPANY = C.COMPANY
   AND IASBAS030.MATGRP = C.MATGRP
   AND IASBAS041.CLIENT = IASBAS030.CLIENT
   AND IASBAS041.COMPANY = IASBAS030.COMPANY
   AND IASBAS041.MAINMATGRP = IASBAS030.MAINMATGRP
 ORDER BY A.COMPANY, A.CENTRAL, A.CUSTOMER, A.VALIDFROM, A.DOCTYPE, A.DOCNUM, B.ITEMNUM;

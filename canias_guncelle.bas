Attribute VB_Name = "CaniasModul"
'==========================================================================
' STOK ve SIPARIS sekmelerindeki "CANIAS'tan Guncelle" butonlarinin makrosu.
'
' Baglanti bilgisi ve sorgular gizli "CANIAS" sayfasindadir; bu sayfayi
' excel_kurulum.py doldurur (stok_sorgu.sql, siparis_sorgu.sql,
' canias_baglanti.json). Sifre hicbir yere kaydedilmez: ilk guncellemede
' SQL Server giris penceresi acilir, Excel kapanana kadar hatirlanir.
'
' Sonuc: 2. satir = basliklar, 3. satirdan itibaren veri. 1. satir butona
' ve "son guncelleme" bilgisine ayrilmistir.
'==========================================================================
Option Explicit

Private Const BASLIK_SATIRI As Long = 2
Private mBaglanti As String          ' basarili baglanti (sadece bu Excel oturumunda)

Public Sub CaniasGuncelle()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    On Error Resume Next
    Set ws = ActiveSheet.Shapes(Application.Caller).Parent
    On Error GoTo 0
    Guncelle ws
End Sub

Private Function Ayar() As Worksheet
    On Error Resume Next
    Set Ayar = ThisWorkbook.Worksheets("CANIAS")
    On Error GoTo 0
End Function

Private Function SorguBul(ws As Worksheet) As String
    Dim a As Worksheet, r As Long
    Set a = Ayar()
    If a Is Nothing Then Exit Function
    For r = 7 To 20
        If CStr(a.Cells(r, 1).Value) = ws.Name Then
            SorguBul = CStr(a.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
End Function

Private Function Baglan() As Object
    Dim cn As Object, a As Worksheet
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 30
    cn.CommandTimeout = 600
    If mBaglanti <> "" Then
        On Error Resume Next
        cn.Open mBaglanti
        If Err.Number = 0 Then
            On Error GoTo 0
            Set Baglan = cn
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
        mBaglanti = ""
        Set cn = CreateObject("ADODB.Connection")
        cn.ConnectionTimeout = 30
        cn.CommandTimeout = 600
    End If
    Set a = Ayar()
    cn.Provider = "MSDASQL"
    cn.Properties("Prompt") = 2          ' adPromptComplete: eksik sifre icin giris penceresi
    cn.Open "Driver={" & a.Range("B4").Value & "};Server=" & a.Range("B1").Value & _
            ";Database=" & a.Range("B2").Value & ";UID=" & a.Range("B3").Value & ";"
    mBaglanti = cn.ConnectionString      ' sifre yalnizca bellekte tutulur
    Set Baglan = cn
End Function

Private Sub Guncelle(ws As Worksheet)
    Dim sql As String, cn As Object, rs As Object
    Dim i As Long, n As Long, sonSatir As Long, sonKolon As Long
    Dim t0 As Double

    If Ayar() Is Nothing Then
        MsgBox "Gizli 'CANIAS' ayar sayfasi bulunamadi. excel_kurulum.py calistirilmali.", vbExclamation
        Exit Sub
    End If
    sql = SorguBul(ws)
    If Len(Trim(sql)) = 0 Then
        MsgBox "'" & ws.Name & "' sayfasi icin CANIAS sayfasinda sorgu yok.", vbExclamation
        Exit Sub
    End If

    t0 = Timer
    Application.StatusBar = "CANIAS'a baglaniliyor..."
    On Error GoTo Hata
    Set cn = Baglan()
    Application.StatusBar = "CANIAS sorgusu calisiyor (" & ws.Name & ")..."
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3                ' adUseClient: RecordCount icin
    rs.Open sql, cn, 3, 1
    n = rs.RecordCount
    If n <= 0 Then
        rs.Close: cn.Close
        Application.StatusBar = False
        MsgBox "Sorgu hic satir dondurmedi. '" & ws.Name & "' sayfasina dokunulmadi.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    sonSatir = ws.Cells.SpecialCells(xlCellTypeLastCell).Row
    sonKolon = ws.Cells.SpecialCells(xlCellTypeLastCell).Column
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    If sonSatir >= BASLIK_SATIRI Then
        ws.Range(ws.Cells(BASLIK_SATIRI, 1), ws.Cells(sonSatir, Application.Max(sonKolon, rs.Fields.Count))).Clear
    End If

    For i = 0 To rs.Fields.Count - 1
        With ws.Cells(BASLIK_SATIRI, i + 1)
            .Value = rs.Fields(i).Name
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 56, 100)
            .WrapText = True
        End With
        Select Case rs.Fields(i).Type
            Case 7, 133, 135               ' tarih alanlari
                ws.Range(ws.Cells(BASLIK_SATIRI + 1, i + 1), ws.Cells(BASLIK_SATIRI + n, i + 1)).NumberFormat = "dd.mm.yyyy"
            Case 2, 3, 4, 5, 6, 14, 17, 20, 131
                ws.Range(ws.Cells(BASLIK_SATIRI + 1, i + 1), ws.Cells(BASLIK_SATIRI + n, i + 1)).NumberFormat = "#,##0.###"
        End Select
    Next i
    sonKolon = rs.Fields.Count
    ws.Cells(BASLIK_SATIRI + 1, 1).CopyFromRecordset rs
    rs.Close
    cn.Close

    ws.Range(ws.Cells(BASLIK_SATIRI, 1), ws.Cells(BASLIK_SATIRI + n, sonKolon)).AutoFilter
    ws.Cells(1, DurumKolonu(ws)).Value = _
        "Son g" & ChrW(252) & "ncelleme: " & Format(Now, "dd.mm.yyyy hh:nn") & "  -  " & n & " sat" & ChrW(305) & "r (CANIAS)"

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    ThisWorkbook.Save
    MsgBox ws.Name & " g" & ChrW(252) & "ncellendi: " & n & " sat" & ChrW(305) & "r, " & _
           Format(Timer - t0, "0.0") & " sn." & vbCrLf & "Dosya kaydedildi.", vbInformation
    Exit Sub

Hata:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    If Not cn Is Nothing Then If cn.State <> 0 Then cn.Close
    mBaglanti = ""
    MsgBox "CANIAS g" & ChrW(252) & "ncellemesi ba" & ChrW(351) & "ar" & ChrW(305) & "s" & ChrW(305) & "z:" & _
           vbCrLf & vbCrLf & Err.Description, vbCritical
End Sub

' Butonun sagindaki ilk sutun ("son guncelleme" yazisi icin)
Private Function DurumKolonu(ws As Worksheet) As Long
    Dim s As Shape, sag As Double, c As Long
    sag = 0
    For Each s In ws.Shapes
        If s.Name = "btnCanias" Then sag = s.Left + s.Width
    Next s
    c = 1
    Do While ws.Cells(1, c).Left < sag + 6 And c < 50
        c = c + 1
    Loop
    DurumKolonu = c
End Function

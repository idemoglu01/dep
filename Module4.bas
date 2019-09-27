Attribute VB_Name = "Module4"
Sub Makro1()
Attribute Makro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro1 Makro
'

'
    Sheets(Array("AnaSayfa", "EGT", "Defibrilatör", "Manometre", "DENEME 4", "Cihazlar", _
        "Cihaz Res_Uzman_Ýmza", "Cihaz Res_baslik", "Cihaz Uzman_Bilgiler", _
        "Cihaz Res_sysii", "Resim Dosyalarý", "Liste", "Lokasyon Hazýrlama", _
        "ElemanListesi", "Önizleme", "Kalibratorler", "Kalibratorler2", "Ateþ Ölçer", _
        "CihazYetkinlik", "Cihaz Listesi", "Cihaz Listesi2", "Aðrý Pompasý", "Alkolmetre", _
        "Anestezi Cihazý", "Aspiratör")).Select
    Sheets("EGT").Activate
    Sheets(Array("Basýnç Manometre", "Bebek Tartýsý", "Beslenme Pompasý", "Buzdolabý", _
        "C Kollu Röntgen", "Coldpack", "Derece", "Defibrilatör AED", "Derin Dondurucu", _
        "Efor Testi", "EKG", "Eko Cihazý", "El Dopleri", "Benmari", "Etüv", "Flash Otoklav", _
        "Fototerapi", "FTR USG", "Glukometre", "Hassas Terazi", "Hastabaþý Monitör", _
        "Holter Ritim", "Holter Tansiyon", "Hotpack Kazaný", "Ýnfüzyon Pompasý")).Select _
        Replace:=False
    Sheets(Array("Ýnkübatör", "Kan Dolabý", "Kombine Elektroterapi", "Koter", "Küvöz", _
        "Küvöz Transport", "Mamografi", "Mikro Pipet", "Mikro Pipet Ayar", "Morg Dolabý", _
        "Nem Ölçer", "Neopuff", "NST", "O2 Flowmetre", "Otoklav", "Pace Maker", _
        "Panoramik Röntgen", "Parafin Kazaný", "Peak Flowmetre", "Periapikal Röntgen", _
        "Pet-Ct", "Puls Oksimetre", "Radyan Isýtýcý", "Röntgen", "Santrifüj")).Select _
        Replace:=False
    Sheets(Array("Seyyar Röntgen", "SFT", "Sýcaklýk Nem Ölçer", "Sývý Azot Tanký", _
        "Þýrýnga Pompasý", "Tansiyon Aleti Manuel", "Tansiyon Aleti Otomatik", "Tartý", _
        "Tens", "Tomografi", "Turnike", "Ultrason", "Vakum Manometre", "Vaporizatör", _
        "Ventilatör", "Ventilatör Transport", "Vital Monitör")).Select Replace:=False
    Sheets(Array("EGT", "Defibrilatör", "Manometre", "DENEME 4", "Cihazlar", _
        "Cihaz Res_Uzman_Ýmza", "Cihaz Res_baslik", "Cihaz Uzman_Bilgiler", _
        "Cihaz Res_sysii", "Resim Dosyalarý", "Liste", "Lokasyon Hazýrlama", _
        "ElemanListesi", "Önizleme", "Kalibratorler", "Kalibratorler2", "Ateþ Ölçer", _
        "CihazYetkinlik", "Cihaz Listesi", "Cihaz Listesi2", "Aðrý Pompasý", "Alkolmetre", _
        "Anestezi Cihazý", "Aspiratör", "Basýnç Manometre")).Select
    Sheets("EGT").Activate
    Sheets(Array("Bebek Tartýsý", "Beslenme Pompasý", "Buzdolabý", "C Kollu Röntgen", _
        "Coldpack", "Derece", "Defibrilatör AED", "Derin Dondurucu", "Efor Testi", "EKG", _
        "Eko Cihazý", "El Dopleri", "Benmari", "Etüv", "Flash Otoklav", "Fototerapi", _
        "FTR USG", "Glukometre", "Hassas Terazi", "Hastabaþý Monitör", "Holter Ritim", _
        "Holter Tansiyon", "Hotpack Kazaný", "Ýnfüzyon Pompasý", "Ýnkübatör")).Select _
        Replace:=False
    Sheets(Array("Kan Dolabý", "Kombine Elektroterapi", "Koter", "Küvöz", _
        "Küvöz Transport", "Mamografi", "Mikro Pipet", "Mikro Pipet Ayar", "Morg Dolabý", _
        "Nem Ölçer", "Neopuff", "NST", "O2 Flowmetre", "Otoklav", "Pace Maker", _
        "Panoramik Röntgen", "Parafin Kazaný", "Peak Flowmetre", "Periapikal Röntgen", _
        "Pet-Ct", "Puls Oksimetre", "Radyan Isýtýcý", "Röntgen", "Santrifüj", _
        "Seyyar Röntgen")).Select Replace:=False
    Sheets(Array("SFT", "Sýcaklýk Nem Ölçer", "Sývý Azot Tanký", "Þýrýnga Pompasý", _
        "Tansiyon Aleti Manuel", "Tansiyon Aleti Otomatik", "Tartý", "Tens", "Tomografi", _
        "Turnike", "Ultrason", "Vakum Manometre", "Vaporizatör", "Ventilatör", _
        "Ventilatör Transport", "Vital Monitör")).Select Replace:=False
    Sheets(Array("EGT", "Defibrilatör", "Manometre", "DENEME 4", "Cihazlar", _
        "Cihaz Res_Uzman_Ýmza", "Cihaz Res_baslik", "Cihaz Uzman_Bilgiler", _
        "Cihaz Res_sysii", "Resim Dosyalarý", "Liste", "Lokasyon Hazýrlama", _
        "ElemanListesi", "Önizleme", "Kalibratorler", "Kalibratorler2", "Ateþ Ölçer", _
        "CihazYetkinlik", "Cihaz Listesi", "Cihaz Listesi2", "Aðrý Pompasý", "Alkolmetre", _
        "Anestezi Cihazý", "Aspiratör", "Basýnç Manometre")).Select
    Sheets("EGT").Activate
    Sheets(Array("Bebek Tartýsý", "Beslenme Pompasý", "Buzdolabý", "C Kollu Röntgen", _
        "Coldpack", "Derece", "Defibrilatör AED", "Derin Dondurucu", "Efor Testi", "EKG", _
        "Eko Cihazý", "El Dopleri", "Benmari", "Etüv", "Flash Otoklav", "Fototerapi", _
        "FTR USG", "Glukometre", "Hassas Terazi", "Hastabaþý Monitör", "Holter Ritim", _
        "Holter Tansiyon", "Hotpack Kazaný", "Ýnfüzyon Pompasý", "Ýnkübatör")).Select _
        Replace:=False
    Sheets(Array("Kan Dolabý", "Kombine Elektroterapi", "Koter", "Küvöz", _
        "Küvöz Transport", "Mamografi", "Mikro Pipet", "Mikro Pipet Ayar", "Morg Dolabý", _
        "Nem Ölçer", "Neopuff", "NST", "O2 Flowmetre", "Otoklav", "Pace Maker", _
        "Panoramik Röntgen", "Parafin Kazaný", "Peak Flowmetre", "Periapikal Röntgen", _
        "Pet-Ct", "Puls Oksimetre", "Radyan Isýtýcý", "Röntgen", "Santrifüj", _
        "Seyyar Röntgen")).Select Replace:=False
    Sheets(Array("SFT", "Sýcaklýk Nem Ölçer", "Sývý Azot Tanký", "Þýrýnga Pompasý", _
        "Tansiyon Aleti Manuel", "Tansiyon Aleti Otomatik", "Tartý", "Tens", "Tomografi", _
        "Turnike", "Ultrason", "Vakum Manometre", "Vaporizatör", "Ventilatör", _
        "Ventilatör Transport", "Vital Monitör")).Select Replace:=False
    ActiveWindow.SelectedSheets.Delete
End Sub
Sub backupIn()
    'Call sheetVH
    If CreateObject("Scripting.FileSystemObject").FileExists("C:\BioCaCe" + Replace(AnaS2.Range("B1"), " ", "") + "Yedek.dll") = False Then Exit Sub
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Workbooks.Open "C:\BioCaCe" + Replace(AnaS2.Range("B1"), " ", "") + "Yedek.dll", Password:="#idem#", ReadOnly:=True
    
    
    Dim backupArray(), backupArray2() As String
    ReDim Preserve backupArray(Workbooks("BioCaCe" + Replace(AnaS2.Range("B1"), " ", "") + "Yedek.dll").Sheets.Count - 1)
    Dim indis As Integer
    indis = 1
    sayac = 0
    For i = 1 To Workbooks("BioCaCe" + Replace(AnaS2.Range("B1"), " ", "") + "Yedek.dll").Sheets.Count
        If sheetControl(Workbooks("BioCaCe" + Replace(AnaS2.Range("B1"), " ", "") + "Yedek.dll").Sheets(i).Name) = True Then
            ReDim Preserve backupArray2(sayac)
            backupArray2(sayac) = Workbooks("BioCaCe" + Replace(AnaS2.Range("B1"), " ", "") + "Yedek.dll").Sheets(i).Name
            sayac = sayac + 1
            ThisWorkbook.Sheets(Workbooks("BioCaCe" + Replace(AnaS2.Range("B1"), " ", "") + "Yedek.dll").Sheets(i).Name).Visible = xlSheetVisible
        End If
        backupArray(i - indis) = Workbooks("BioCaCe" + Replace(AnaS2.Range("B1"), " ", "") + "Yedek.dll").Sheets(i).Name
    Next i
    If sayac <> 0 Then
        ThisWorkbook.Sheets(backupArray2()).Delete
    End If
    Workbooks("BioCaCe" + Replace(AnaS2.Range("B1"), " ", "") + "Yedek.dll").Sheets(backupArray()).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Workbooks("BioCaCe" + Replace(AnaS2.Range("B1"), " ", "") + "Yedek.dll").Close
    Application.DisplayAlerts = True
    
End Sub
Private Function sheetControl(SheetName As String) As Boolean
      For Each WS In ThisWorkbook.Worksheets
        If SheetName = WS.Name Then
          sheetControl = True
          Exit Function
        End If
      Next WS
      sheetControl = False
End Function


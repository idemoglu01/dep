Attribute VB_Name = "Module4"
Sub Makro1()
Attribute Makro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Makro1 Makro
'

'
    Sheets(Array("AnaSayfa", "EGT", "Defibrilat�r", "Manometre", "DENEME 4", "Cihazlar", _
        "Cihaz Res_Uzman_�mza", "Cihaz Res_baslik", "Cihaz Uzman_Bilgiler", _
        "Cihaz Res_sysii", "Resim Dosyalar�", "Liste", "Lokasyon Haz�rlama", _
        "ElemanListesi", "�nizleme", "Kalibratorler", "Kalibratorler2", "Ate� �l�er", _
        "CihazYetkinlik", "Cihaz Listesi", "Cihaz Listesi2", "A�r� Pompas�", "Alkolmetre", _
        "Anestezi Cihaz�", "Aspirat�r")).Select
    Sheets("EGT").Activate
    Sheets(Array("Bas�n� Manometre", "Bebek Tart�s�", "Beslenme Pompas�", "Buzdolab�", _
        "C Kollu R�ntgen", "Coldpack", "Derece", "Defibrilat�r AED", "Derin Dondurucu", _
        "Efor Testi", "EKG", "Eko Cihaz�", "El Dopleri", "Benmari", "Et�v", "Flash Otoklav", _
        "Fototerapi", "FTR USG", "Glukometre", "Hassas Terazi", "Hastaba�� Monit�r", _
        "Holter Ritim", "Holter Tansiyon", "Hotpack Kazan�", "�nf�zyon Pompas�")).Select _
        Replace:=False
    Sheets(Array("�nk�bat�r", "Kan Dolab�", "Kombine Elektroterapi", "Koter", "K�v�z", _
        "K�v�z Transport", "Mamografi", "Mikro Pipet", "Mikro Pipet Ayar", "Morg Dolab�", _
        "Nem �l�er", "Neopuff", "NST", "O2 Flowmetre", "Otoklav", "Pace Maker", _
        "Panoramik R�ntgen", "Parafin Kazan�", "Peak Flowmetre", "Periapikal R�ntgen", _
        "Pet-Ct", "Puls Oksimetre", "Radyan Is�t�c�", "R�ntgen", "Santrif�j")).Select _
        Replace:=False
    Sheets(Array("Seyyar R�ntgen", "SFT", "S�cakl�k Nem �l�er", "S�v� Azot Tank�", _
        "��r�nga Pompas�", "Tansiyon Aleti Manuel", "Tansiyon Aleti Otomatik", "Tart�", _
        "Tens", "Tomografi", "Turnike", "Ultrason", "Vakum Manometre", "Vaporizat�r", _
        "Ventilat�r", "Ventilat�r Transport", "Vital Monit�r")).Select Replace:=False
    Sheets(Array("EGT", "Defibrilat�r", "Manometre", "DENEME 4", "Cihazlar", _
        "Cihaz Res_Uzman_�mza", "Cihaz Res_baslik", "Cihaz Uzman_Bilgiler", _
        "Cihaz Res_sysii", "Resim Dosyalar�", "Liste", "Lokasyon Haz�rlama", _
        "ElemanListesi", "�nizleme", "Kalibratorler", "Kalibratorler2", "Ate� �l�er", _
        "CihazYetkinlik", "Cihaz Listesi", "Cihaz Listesi2", "A�r� Pompas�", "Alkolmetre", _
        "Anestezi Cihaz�", "Aspirat�r", "Bas�n� Manometre")).Select
    Sheets("EGT").Activate
    Sheets(Array("Bebek Tart�s�", "Beslenme Pompas�", "Buzdolab�", "C Kollu R�ntgen", _
        "Coldpack", "Derece", "Defibrilat�r AED", "Derin Dondurucu", "Efor Testi", "EKG", _
        "Eko Cihaz�", "El Dopleri", "Benmari", "Et�v", "Flash Otoklav", "Fototerapi", _
        "FTR USG", "Glukometre", "Hassas Terazi", "Hastaba�� Monit�r", "Holter Ritim", _
        "Holter Tansiyon", "Hotpack Kazan�", "�nf�zyon Pompas�", "�nk�bat�r")).Select _
        Replace:=False
    Sheets(Array("Kan Dolab�", "Kombine Elektroterapi", "Koter", "K�v�z", _
        "K�v�z Transport", "Mamografi", "Mikro Pipet", "Mikro Pipet Ayar", "Morg Dolab�", _
        "Nem �l�er", "Neopuff", "NST", "O2 Flowmetre", "Otoklav", "Pace Maker", _
        "Panoramik R�ntgen", "Parafin Kazan�", "Peak Flowmetre", "Periapikal R�ntgen", _
        "Pet-Ct", "Puls Oksimetre", "Radyan Is�t�c�", "R�ntgen", "Santrif�j", _
        "Seyyar R�ntgen")).Select Replace:=False
    Sheets(Array("SFT", "S�cakl�k Nem �l�er", "S�v� Azot Tank�", "��r�nga Pompas�", _
        "Tansiyon Aleti Manuel", "Tansiyon Aleti Otomatik", "Tart�", "Tens", "Tomografi", _
        "Turnike", "Ultrason", "Vakum Manometre", "Vaporizat�r", "Ventilat�r", _
        "Ventilat�r Transport", "Vital Monit�r")).Select Replace:=False
    Sheets(Array("EGT", "Defibrilat�r", "Manometre", "DENEME 4", "Cihazlar", _
        "Cihaz Res_Uzman_�mza", "Cihaz Res_baslik", "Cihaz Uzman_Bilgiler", _
        "Cihaz Res_sysii", "Resim Dosyalar�", "Liste", "Lokasyon Haz�rlama", _
        "ElemanListesi", "�nizleme", "Kalibratorler", "Kalibratorler2", "Ate� �l�er", _
        "CihazYetkinlik", "Cihaz Listesi", "Cihaz Listesi2", "A�r� Pompas�", "Alkolmetre", _
        "Anestezi Cihaz�", "Aspirat�r", "Bas�n� Manometre")).Select
    Sheets("EGT").Activate
    Sheets(Array("Bebek Tart�s�", "Beslenme Pompas�", "Buzdolab�", "C Kollu R�ntgen", _
        "Coldpack", "Derece", "Defibrilat�r AED", "Derin Dondurucu", "Efor Testi", "EKG", _
        "Eko Cihaz�", "El Dopleri", "Benmari", "Et�v", "Flash Otoklav", "Fototerapi", _
        "FTR USG", "Glukometre", "Hassas Terazi", "Hastaba�� Monit�r", "Holter Ritim", _
        "Holter Tansiyon", "Hotpack Kazan�", "�nf�zyon Pompas�", "�nk�bat�r")).Select _
        Replace:=False
    Sheets(Array("Kan Dolab�", "Kombine Elektroterapi", "Koter", "K�v�z", _
        "K�v�z Transport", "Mamografi", "Mikro Pipet", "Mikro Pipet Ayar", "Morg Dolab�", _
        "Nem �l�er", "Neopuff", "NST", "O2 Flowmetre", "Otoklav", "Pace Maker", _
        "Panoramik R�ntgen", "Parafin Kazan�", "Peak Flowmetre", "Periapikal R�ntgen", _
        "Pet-Ct", "Puls Oksimetre", "Radyan Is�t�c�", "R�ntgen", "Santrif�j", _
        "Seyyar R�ntgen")).Select Replace:=False
    Sheets(Array("SFT", "S�cakl�k Nem �l�er", "S�v� Azot Tank�", "��r�nga Pompas�", _
        "Tansiyon Aleti Manuel", "Tansiyon Aleti Otomatik", "Tart�", "Tens", "Tomografi", _
        "Turnike", "Ultrason", "Vakum Manometre", "Vaporizat�r", "Ventilat�r", _
        "Ventilat�r Transport", "Vital Monit�r")).Select Replace:=False
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


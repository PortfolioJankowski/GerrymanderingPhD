Attribute VB_Name = "Inicjalizacja"
Option Explicit
Public kolekcja_komitetow As Collection
Public kolekcja_powiatow As Collection
Public kolekcja_okregow As Collection



Public Sub RozpocznijProgram()

    Dim wyniki_komitetow As Scripting.Dictionary
    Dim tempDict As Scripting.Dictionary
    Dim nowy_komitet As komitet
    Dim nowy_powiat As Powiat
    Dim nowy_okreg As okreg
    Dim wb As Workbook
    Dim wsK As Worksheet
    Dim wsO As Worksheet
    Dim wsP As Worksheet
    Dim i As Long
    Dim j As Long
    Dim lrK As Long
    Dim lrP As Long
    Dim lrO As Long
    Dim objK As komitet
    Dim tempColl As Collection
    Dim objP As Powiat
    Dim tempMieszk As Long
    Dim sArr As Variant
    Dim key As Variant
    Dim suma_mandatow As Long
    Dim liczba_mandatow As Long
    
    Set wb = ThisWorkbook
    Set wsK = wb.Worksheets("Komitety")
    Set wsO = wb.Worksheets("Okregi")
    Set wsP = wb.Worksheets("Powiaty")
    Set kolekcja_komitetow = New Collection
    Set kolekcja_powiatow = New Collection
    Set kolekcja_okregow = New Collection
    
    
    lrK = wsK.Cells(wsK.Rows.Count, "A").End(xlUp).Row
    lrP = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    lrO = wsO.Cells(wsO.Rows.Count, "A").End(xlUp).Row
    
    'tworze tablice komitetow, do ktorej dodaje obiekty typu komitet
    For i = 2 To lrK
        Set nowy_komitet = New komitet
        nowy_komitet.Initialize wsK.Cells(i, "B").value, wsK.Cells(i, "C").value, wsK.Cells(i, "D").value
        kolekcja_komitetow.Add nowy_komitet
    Next i

    'tworze powiaty
    For i = 2 To lrP
      Set nowy_powiat = New Powiat
      Set wyniki_komitetow = New Scripting.Dictionary
      For j = 13 To 40
        wyniki_komitetow.Add wsP.Cells(1, j).value, wsP.Cells(i, j).value
      Next j
        nowy_powiat.Initialize wsP.Cells(i, "B").value, wsP.Cells(i, "A").value, wsP.Cells(i, "L").value, _
                            Array(wsP.Cells(i, "C").value, wsP.Cells(i, "D").value, wsP.Cells(i, "E").value, wsP.Cells(i, "F").value, wsP.Cells(i, "G").value, wsP.Cells(i, "H").value, wsP.Cells(i, "I").value, wsP.Cells(i, "J").value, wsP.Cells(i, "K").value), 0, wyniki_komitetow
        kolekcja_powiatow.Add nowy_powiat
    Next i
    
    
    'tworze okregi wyborcze (w zaleznosci od wlasciwosci powiatu powiat bedzie wskakiwal do wlasciwosci [kolekcji] odpowiedniego okregu
    For i = 2 To lrO
        Set nowy_okreg = New okreg
        'tworze zmienne tymczasowe, do których laduje powiaty, wyniki, mieszkancow
        '(nr As Long, pColl As Collection, dict As Scripting.Dictionary, m As Long, sArr As Variant)
        Set tempColl = New Collection
        Set tempDict = New Scripting.Dictionary
        tempMieszk = 0
        For Each objP In kolekcja_powiatow
            'jezeli powiat nalezy do tego okregu:
            If objP.okreg = wsO.Cells(i, "A").value Then
                tempColl.Add objP
                tempMieszk = tempMieszk + objP.Mieszkancy
                For Each key In objP.wyniki.keys
                    If tempDict.Exists(key) Then
                        tempDict(key) = tempDict(key) + objP.wyniki(key)
                    Else
                        tempDict.Add key, objP.wyniki(key)
                    End If
                Next key
            End If
        Next objP
                sArr = Array(wsO.Cells(i, "B").value, wsO.Cells(i, "C").value, wsO.Cells(i, "D").value, wsO.Cells(i, "E").value)
                'nr As Long, pColl As Collection, dict As Scripting.Dictionary, m As Long, sArr As Variant
                nowy_okreg.Initialize wsO.Cells(i, "A").value, tempColl, tempDict, tempMieszk, sArr
                
                nowy_okreg.mandaty = Formalnosci.ObliczMandatyOkregu(nowy_okreg)
                kolekcja_okregow.Add nowy_okreg
                
                'czyszcze zmienne tymczasowe
                
                tempMieszk = 0
    'kolejny okreg
    Next i
    
    'obliczam pierwotna liczbe mandatow w calym wojewodztwie
    liczba_mandatow = Formalnosci.ObliczMandatyWojewodztwa
    
    'usuniecie nadwyzkowego mandatu albo dodanie gdy jest ich za malo (jezeli jest), ustawienie normy przedstawicielskiej
    If liczba_mandatow <> MANDATYSLASK Then
        Formalnosci.UsunMandatyNadwyzkowe liczba_mandatow
    End If

End Sub



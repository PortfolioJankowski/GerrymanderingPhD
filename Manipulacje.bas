Attribute VB_Name = "Manipulacje"
Option Explicit

Public mandatyOkregu As Long
Public mandatySasiada As Long
Public wynikiOkregu As Scripting.Dictionary
Public wynikiSasiada As Scripting.Dictionary

Sub RozpocznijPrzesuwanie()
    Dim objO As okreg
    Dim okrSasiad As okreg
    Dim objP As Powiat
    Dim powSasiad As Powiat
    Dim i As Variant
    Dim liczbaSasiadow As Long
    Dim tablicaSasiadow As Variant
    Dim key As Variant
    Dim sprMandaty As Variant
    
   'SCHEMAT DZIALANIA
    'pobieram okreg oraz jego sasiadow
    For Each objO In kolekcja_okregow
        tablicaSasiadow = objO.Sasiedzi
        
        'ide przez sasiadow danego okregu
        For i = LBound(tablicaSasiadow) To UBound(tablicaSasiadow)
                Set okrSasiad = PobierzOkregPoNumerze(CInt(objO.Sasiedzi(i)))
            
            'inicjalizacja ZMIENNYCH GLOBALNYCH
            Set wynikiOkregu = New Dictionary
            Set wynikiSasiada = New Dictionary
            
            'jesli sasiad jest pusty to bierzemy kolejny okreg
            If okrSasiad Is Nothing Then
                GoTo kolejny_okreg
            End If
            
            'przypisanie pierwotnych wyników okregu1 do zmiennej globalnej
            For Each objP In objO.Powiaty
               For Each key In objP.wyniki.keys
                    If wynikiOkregu.Exists(key) Then
                        wynikiOkregu(key) = wynikiOkregu(key) + objP.wyniki(key)
                    Else
                        wynikiOkregu.Add key, objP.wyniki(key)
                    End If
                Next key
            Next objP
            
            'przypisanie pierwotnych wyników okregu2 do zmiennej globalnej
            For Each objP In okrSasiad.Powiaty
               For Each key In objP.wyniki.keys
                    If wynikiSasiada.Exists(key) Then
                        wynikiSasiada(key) = wynikiSasiada(key) + objP.wyniki(key)
                    Else
                        wynikiSasiada.Add key, objP.wyniki(key)
                    End If
                Next key
            Next objP
                 
                mandatyOkregu = Formalnosci.dHondt(wynikiOkregu, objO.Numer)
                mandatySasiada = Formalnosci.dHondt(wynikiSasiada, okrSasiad.Numer)
                'ROZPOCZYNAM PRZENOSZENIE POWIATÓW
                PrzeniesPowiaty objO, okrSasiad
        Next i
kolejny_okreg:
    Next objO
    
    Debug.Print "XD"
  
    
    
End Sub

Function PobierzOkregPoNumerze(nr As Integer) As okreg
    Dim objO As okreg
    Dim objP As Powiat
    
    For Each objO In kolekcja_okregow
        If objO.Numer = nr Then
            Set PobierzOkregPoNumerze = objO
            Exit Function
        Else
            Set PobierzOkregPoNumerze = Nothing
        End If
    Next objO
End Function

Sub PrzeniesPowiaty(okreg As okreg, sasiad As okreg)
    Dim objP As Powiat
    Dim powSasiad As Powiat
    Dim indexDoUsuniecia As Long
    Dim i As Long
    Dim j As Long
    Dim nazwa As String
    
For i = okreg.Powiaty.Count To 1 Step -1
    'pobieram nazwe aby potem usunac dobry powiat
    Set objP = okreg.Powiaty(i)
    nazwa = objP.NazwaPowiatu
    
    For Each powSasiad In sasiad.Powiaty
        If PowiatySasiaduja(objP, powSasiad) Then
            ' PRZESUNIECIE jednego
            sasiad.Powiaty.Add objP
            
            ' Find the index of the object in okreg.Powiaty and remove it
            Dim indexToRemove As Long
            For j = 1 To okreg.Powiaty.Count
                If okreg.Powiaty(j) Is objP Then
                    indexToRemove = j
                    Exit For
                End If
            Next j
            
            If indexToRemove > 0 Then
                okreg.Powiaty.Remove indexToRemove
            End If

            If Not OsiagnietoLepszyWynik(okreg.Numer, sasiad.Numer) Then
                ' Reverse the changes if the condition is not met
                okreg.Powiaty.Add objP
                
                ' Find the index of the object in sasiad.Powiaty and remove it
                Dim foundIndex As Long
                For j = 1 To sasiad.Powiaty.Count
                    If sasiad.Powiaty(j).NazwaPowiatu = nazwa Then
                        foundIndex = j
                        Exit For
                    End If
                Next j
                
                If foundIndex > 0 Then
                    sasiad.Powiaty.Remove foundIndex
                End If
            End If
        End If
    Next powSasiad
Next i


End Sub

Function PowiatySasiaduja(powiat1 As Powiat, powiat2 As Powiat) As Boolean
    Dim nazwa_sasiada As Variant
    
    For Each nazwa_sasiada In powiat1.Sasiedzi
        If IsEmpty(nazwa_sasiada) Then
            PowiatySasiaduja = False
            Exit Function
        End If
        If nazwa_sasiada = powiat2.NazwaPowiatu Then
            PowiatySasiaduja = True
            Exit Function
        End If
    Next nazwa_sasiada
            PowiatySasiaduja = False
End Function

Function OsiagnietoLepszyWynik(nr As Long, nr2 As Long) As Boolean
    Dim objO As okreg
    Dim objP As Powiat
    Dim wyniki As Scripting.Dictionary
    Dim wyniki2 As Scripting.Dictionary
    Dim mandaty As Long
    Dim mandaty2 As Long
    Dim key As Variant
    
    Set wyniki = New Dictionary
    Set wyniki2 = New Dictionary
    
    
    'dodaje do tymczasowych slownikow dla obu okregow wyniki ze wszystkich powiatów
    For Each objO In kolekcja_okregow
        'sprawdzam czy ten okreg ma nr taki jak pobralem,
        If objO.Numer = nr Then
            For Each objP In objO.Powiaty
               For Each key In objP.wyniki.keys
                    If wyniki.Exists(key) Then
                        wyniki(key) = wyniki(key) + objP.wyniki(key)
                    Else
                        wyniki.Add key, objP.wyniki(key)
                    End If
                Next key
            Next objP
        End If
        If objO.Numer = nr2 Then
            For Each objP In objO.Powiaty
               For Each key In objP.wyniki.keys
                    If wyniki2.Exists(key) Then
                        wyniki2(key) = wyniki2(key) + objP.wyniki(key)
                    Else
                        wyniki2.Add key, objP.wyniki(key)
                    End If
                Next key
            Next objP
        End If
    Next objO
    
    'oblicznenie glosow na mandaty formula dhondta
    mandaty = Formalnosci.dHondt(wyniki, nr)
    mandaty2 = Formalnosci.dHondt(wyniki2, nr2)
    
    'sprawdzenie mandatów z mandatami na poczatku mieszania
    If mandaty + mandaty2 > mandatyOkregu + mandatySasiada Then
        'jezeli jest lepszy wynik i zachowane kodeksowe limity dla obu okregow
        If Formalnosci.SprawdzKodeksoweLimity(mandaty, mandaty2) Then
            OsiagnietoLepszyWynik = True
        End If
    Else
        OsiagnietoLepszyWynik = False
    End If
    'sprawdzenie czy ilosc mandatow sie zmienila
    'jezeli osiagnieto lepszy wynik to sprawdzam kodeksowe limity
    'to przypisuje do globalnej zmiennej nowy nowy wynik glosowania
    
'mam publiczny slownik nr okregu : mandaty
'sprawdzam czy w nowych okregach jest lepiej niz bylo
'tu trzeba bedzie metode dHondta xD
End Function



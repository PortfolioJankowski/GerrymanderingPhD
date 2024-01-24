Attribute VB_Name = "Formalnosci"
Option Explicit
Public Const JEDNOLITANORMAPRZEDSTAWICIELSTWA As Double = 95785.444444444
Public Const MANDATYSLASK As Integer = 45

Function ObliczMandatyOkregu(okreg As okreg) As Long
    Dim output As Long
    output = Round(okreg.Mieszkancy / JEDNOLITANORMAPRZEDSTAWICIELSTWA)
    ObliczMandatyOkregu = output
End Function

Function ObliczMandatyWojewodztwa() As Long
    Dim objO As okreg
    Dim temp As Long
    temp = 0
    For Each objO In kolekcja_okregow
        temp = temp + objO.mandaty
    Next objO
    
    ObliczMandatyWojewodztwa = temp
End Function

Sub UsunMandatyNadwyzkowe(mandaty As Long)
'419 par 2 pkt 2 -> jezeli jest nadwyzkowy mandat, to nalezy odjac go od mandatow z okregu gdzie norma przedstawicielska jest najmniejsza
'jezeli mandatow jest za malo to dodaje go tam gdzie jest najwyzsza okregowa norma przedstawicielstwa

    Dim objO As okreg
    Dim normy_przedstawicielstwa As Scripting.Dictionary
    Dim min_norma As Double
    Dim max_norma As Double
    Set normy_przedstawicielstwa = New Dictionary
    min_norma = 9999999
    max_norma = 0
    Dim nr_okregu As Integer
    
    While mandaty <> MANDATYSLASK
    
        For Each objO In kolekcja_okregow
                objO.NormaPrzedstawicielska = objO.Mieszkancy / objO.mandaty
                normy_przedstawicielstwa.Add objO.NormaPrzedstawicielska, objO.Numer
                If objO.NormaPrzedstawicielska < min_norma Then
                    min_norma = objO.NormaPrzedstawicielska
                End If
                If objO.NormaPrzedstawicielska > max_norma Then
                    max_norma = objO.NormaPrzedstawicielska
                End If
        Next objO
        
        If mandaty > MANDATYSLASK Then
         'szukam najmniejszej normy przedstawicielstwa w okregu i od niego odejmuj
            nr_okregu = normy_przedstawicielstwa(min_norma)
            For Each objO In kolekcja_okregow
                If nr_okregu = objO.Numer Then
                    objO.mandaty = objO.mandaty - 1
                    mandaty = mandaty - 1
                    GoTo nastepny_mandat
                End If
            Next objO
        Else
        'szukam najwiekszej normy przedstawicielstwa w okregu i do niego dodaje
        nr_okregu = normy_przedstawicielstwa(max_norma)
            For Each objO In kolekcja_okregow
                If nr_okregu = objO.Numer Then
                    objO.mandaty = objO.mandaty + 1
                    mandaty = mandaty + 1
                    GoTo nastepny_mandat
                End If
            Next objO
        End If
        
nastepny_mandat:
    min_norma = 0
    max_norma = 0
    
    Wend
End Sub


Function SprawdzKodeksoweLimity(mandaty1 As Long, mandaty2 As Long) As Boolean
'art 463 w okregu wyborczym wybiera sie od 5 do 15 radnych

    If mandaty1 >= 5 And mandaty1 <= 15 And mandaty2 >= 5 And mandaty2 <= 15 Then
        SprawdzKodeksoweLimity = True
    Else
        SprawdzKodeksoweLimity = False
    End If
    
End Function

Function dHondt(votes As Dictionary, nr As Long) As Long
    Dim results As New Dictionary
    Dim keys() As Variant
    Dim i As Integer
    Dim j As Integer
    Dim mandatyOkregu As Long
    Dim objO As okreg
    Dim mandatyWojewodztwa As Long
    
    ' Extract keys from the input dictionary
    keys = votes.keys
    
    ' Initialize results dictionary
    For i = LBound(keys) To UBound(keys)
        results(keys(i)) = 0
    Next i
    
    For Each objO In kolekcja_okregow
        If objO.Numer = nr Then
            mandatyOkregu = ObliczMandatyOkregu(objO)
            GoTo dalej
        End If
    Next objO
dalej:
    'weryfikuje liczbe mandatow i usuwam nadwyzkowe
    mandatyWojewodztwa = ObliczMandatyWojewodztwa
    UsunMandatyNadwyzkowe mandatyWojewodztwa
    
    'po sprawdzeniu mandatów nadwyzkowych przypisuje jeszcze raz mandaty do okregu
    For Each objO In kolekcja_okregow
        If objO.Numer = nr Then
            mandatyOkregu = objO.mandaty
            Exit For
        End If
    Next objO
    
    
    ' Calculate seats for each party using D'Hondt method
    For i = 1 To mandatyOkregu
        Dim maxKey As Variant
        Dim maxVotes As Long
        maxVotes = 0
        
        ' Find the party with the highest current quotient
        For j = LBound(keys) To UBound(keys)
            Dim currentVotes As Long
            currentVotes = votes(keys(j)) / (results(keys(j)) + 1)
            
            If currentVotes > maxVotes Then
                maxVotes = currentVotes
                maxKey = keys(j)
            End If
        Next j
        
        ' Allocate a seat to the party with the highest quotient
        results(maxKey) = results(maxKey) + 1
    Next i
    
    ' Return the results dictionary
    dHondt = results(wybrany_komitet)
End Function


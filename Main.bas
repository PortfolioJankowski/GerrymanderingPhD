Attribute VB_Name = "Main"
Public wybrany_komitet As String

Public Sub main()

    Inicjalizacja.RozpocznijProgram
    Manipulacje.RozpocznijPrzesuwanie
    
    
'SCHEMAT DZIALANIA
    'Wybieram ugrupowanie, dla ktorego przygotowuje najlepszy scenariusz okregow wyborczych
    'Zalozenie => ide przez kazdy okreg wyborczy, a w kazdym okregu ide od 1 do ostatniego powiatu
    'Zamiana powiatu miedzy okregami:
        'przeskoczenie miedzy kolekcjami powiatu, a co za tym idzie
            'obliczenie na nowo liczby mandatow w nowoutworzonych okregach (aktualizacja wlasciwosci okregow)
                'weryfikacja z kodeksem wyborczym
            'obliczenie na nowo ilosci glosow oddanych na ugrupowania w okregach
    'Weryfikacja stanu mandatów:
        'przeliczenie glosow na mandaty algorytmem d'Hondta
        'Porownanie ilosci mandatow dla wybranej partii -> stan pierwotny a nowy (jak sie zmienilo na + to mieszaj kolejny powiat)
    
'OGRANICZENIA:
'Mozna dodac wlasciwosc do powiatu - liczba przesuniec (np. mozna powiat przesunac max 5x zeby program nie mial nieskonczonej petli)

End Sub


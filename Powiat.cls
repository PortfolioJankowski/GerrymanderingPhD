VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Powiat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pNazwaPowiatu As String
Private pOkreg As Long
Private pMieszkancy As Long
Private pSasiedzi As Variant
Private pPrzesuniecie As Long
Private pWyniki As Scripting.Dictionary

Public Property Get NazwaPowiatu() As String
    NazwaPowiatu = pNazwaPowiatu
End Property
Public Property Let NazwaPowiatu(ByVal RHS As String)
    pNazwaPowiatu = RHS
End Property
Public Property Get okreg() As Long
    okreg = pOkreg
End Property
Public Property Let okreg(ByVal RHS As Long)
    pOkreg = RHS
End Property
Public Property Let Mieszkancy(ByVal value As Long)
    pMieszkancy = value
End Property
Public Property Get Mieszkancy() As Long
    Mieszkancy = pMieszkancy
End Property
Public Property Let Sasiedzi(ByVal value As Variant)
    pSasiedzi = value
End Property
Public Property Get Sasiedzi() As Variant
    Sasiedzi = pSasiedzi
End Property
Public Property Let Przesuniencie(ByVal value As Long)
    pPrzesuniecie = value
End Property
Public Property Get Przesuniencie() As Long
   Przesuniencie = pPrzesuniecie
End Property

Public Property Get wyniki() As Scripting.Dictionary
    If pWyniki Is Nothing Then
        Set wyniki = New Scripting.Dictionary
    Else
        Set wyniki = pWyniki
    End If
End Property
Public Property Set wyniki(ByVal dict As Scripting.Dictionary)
    Set pWyniki = dict
End Property


' Konstruktor klasy
Public Sub Initialize(ByVal NazwaPowiatu As String, ByVal nrOkregu As Long, ByVal LiczbaMieszkancow As Long, ByVal Sasiedzi As Variant, ByVal liczbaPrzesuniec As Long, ByVal daneUgrupowan As Scripting.Dictionary)
    Me.NazwaPowiatu = NazwaPowiatu
    Me.okreg = nrOkregu
    Me.Mieszkancy = LiczbaMieszkancow
    Me.Sasiedzi = Sasiedzi
    Me.Przesuniencie = liczbaPrzesuniec
    Set Me.wyniki = daneUgrupowan
End Sub


Public Function pobierzSasiada(index As Long) As Long
        If IsArray(Sasiedzi) And index >= LBound(Sasiedzi) And index <= UBound(Sasiedzi) Then
            pobierzSasiada = Sasiedzi(index)
        Else
            pobierzSasiada = 99 ' Zwracamy 0, jeœli podany index jest poza zakresem tablicy
        End If
    End Function


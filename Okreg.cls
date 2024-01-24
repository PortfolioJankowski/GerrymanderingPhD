VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Okreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pNumer As Long
Private pSasiedzi As Variant
Private pPowiaty As Collection
Private pMieszkancy As Long
Private pKomitety As New Scripting.Dictionary
Private pMandaty As Long
Private pNormaPrzedstawicielska As Double

Public Property Get Numer() As Long
    Numer = pNumer
End Property
Public Property Let Numer(value As Long)
    pNumer = value
End Property
Public Property Get Sasiedzi() As Variant
    Sasiedzi = pSasiedzi
End Property
Public Property Let Sasiedzi(value As Variant)
    pSasiedzi = value
End Property
Public Property Get Powiaty() As Collection
    If pPowiaty Is Nothing Then
        Set Powiaty = New Collection
    Else
        Set Powiaty = pPowiaty
    End If
End Property
Public Property Set Powiaty(value As Collection)
    Set pPowiaty = value
End Property
Public Property Get Mieszkancy() As Long
    Mieszkancy = pMieszkancy
End Property
Public Property Let Mieszkancy(value As Long)
    pMieszkancy = value
End Property
Public Property Get Komitety() As Scripting.Dictionary
    If pKomitety Is Nothing Then
        Set Komitety = New Scripting.Dictionary
    Else
        Set Komitety = pKomitety
    End If
End Property
Public Property Set Komitety(value As Scripting.Dictionary)
    Set pKomitety = value
End Property
Public Property Get mandaty() As Long
    mandaty = pMandaty
End Property
Public Property Let mandaty(value As Long)
    pMandaty = value
End Property
Public Property Get NormaPrzedstawicielska() As Double
    NormaPrzedstawicielska = pNormaPrzedstawicielska
End Property
Public Property Let NormaPrzedstawicielska(value As Double)
    pNormaPrzedstawicielska = value
End Property

Public Sub Initialize(nr As Long, pColl As Collection, dict As Scripting.Dictionary, m As Long, sArr As Variant)
    Me.Numer = nr
    Set Me.Powiaty = pColl
    Set Me.Komitety = dict
    Me.Mieszkancy = m
    Me.Sasiedzi = sArr
    Me.mandaty = 0
    Me.NormaPrzedstawicielska = 0
End Sub

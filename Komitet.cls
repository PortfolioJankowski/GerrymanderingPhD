VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Komitet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pNazwaKomitetu As String
Private pGlosy As Long
Private pMandaty As Long

Public Property Get NazwaKomitetu() As String
    NazwaKomitetu = pNazwaKomitetu
End Property
Public Property Let NazwaKomitetu(ByVal RHS As String)
    pNazwaKomitetu = RHS
End Property
Public Property Get Glosy() As Long
    Glosy = pGlosy
End Property
Public Property Let Glosy(ByVal RHS As Long)
    pGlosy = RHS
End Property
Public Property Get mandaty() As Long
    mandaty = pMandaty
End Property
Public Property Let mandaty(ByVal RHS As Long)
    pMandaty = RHS
End Property


Public Sub Initialize(ByVal NazwaKomitetu As String, ByVal Glosy As Long, ByVal mandaty As Long)
    Me.NazwaKomitetu = NazwaKomitetu
    Me.Glosy = Glosy
    Me.mandaty = mandaty
End Sub


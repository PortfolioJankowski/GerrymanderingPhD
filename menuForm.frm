VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} menuForm 
   Caption         =   "Menu"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5355
   OleObjectBlob   =   "menuForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "menuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    main.main
End Sub

Private Sub kwCmb_Change()
    wybrany_komitet = kwCmb.value
End Sub

Private Sub UserForm_Initialize()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lr As Long
    Dim i As Long
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Start")
    lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
        
    For i = 1 To lr
        kwCmb.AddItem ws.Cells(i, "A").value
    Next i
    
    
End Sub

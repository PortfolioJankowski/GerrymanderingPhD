Attribute VB_Name = "Wizualizacje"
Option Explicit

Public Sub wizualizuj()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shape As shape
    
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Start")
    
    Set shape = ws.Shapes("zawierciañski")
    shape.Glow.Color.RGB = RGB(250, 0, 0)
    
End Sub

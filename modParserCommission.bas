Attribute VB_Name = "modParserCommission"

Option Explicit

Public Function ParseCommissionTable() As Object
    Dim ws As Worksheet
    Dim dict As Object
    Dim lastRow As Long
    Dim r As Long
    
    Dim idComer As String
    Dim comm As clsCommission
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets("Comisioane") 
    
    ' Find last used row in column A (Id Terminal)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Start from row 2 (skip headers)
    For r = 2 To lastRow
        
        idComer = Trim(ws.Cells(r, "A").Value)  
            
        If idComer <> "" Then
            Set comm = New clsCommission
            With comm
                .CommissionPercent = CDbl(ws.Cells(r, "B").Value)
                .MinCommission = CDbl(ws.Cells(r, "C").Value)
                .MaxCommission = CDbl(ws.Cells(r, "D").Value)
            End With
            
            Set dict(idComer) = comm
        End If
        
    Next r
    
    Set ParseCommissionTable = dict
End Function

Attribute VB_Name = "modParserCommission"
Option Explicit

Public Function ParseCommissionTable() As Object
    Dim ws As Worksheet
    Dim dict As Object
    Dim lastRow As Long
    Dim r As Long
    
    Dim idTerm As String
    Dim comm As clsCommission
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets("Comisioane")
    
    ' Validate commission table first
    If Not ValidateCommissionTable(ws) Then
        Set ParseCommissionTable = Nothing
        Exit Function
    End If
    
    ' Find last used row in column A (Id Terminal)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Start from row 2 (skip headers)
    For r = 2 To lastRow
        
        idTerm = Trim(ws.Cells(r, "A").Value)
            
        If idTerm <> "" Then
            Set comm = New clsCommission
            With comm
                .CommissionPercent = CDbl(ws.Cells(r, "B").Value)
                .MinCommission = CDbl(ws.Cells(r, "C").Value)
                .MaxCommission = CDbl(ws.Cells(r, "D").Value)
            End With
            
            Set dict(idTerm) = comm
        End If
        
    Next r
    
    Set ParseCommissionTable = dict
End Function

Private Function ValidateCommissionTable(ws As Worksheet) As Boolean
    Dim lastRow As Long
    Dim r As Long
    Dim idTerm As String
    Dim errorMsg As String
    
    ValidateCommissionTable = True
    errorMsg = ""
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For r = 2 To lastRow
        idTerm = Trim(ws.Cells(r, "A").Value)
        
        If idTerm <> "" Then
            ' Only check for missing commission percent (required field)
            If IsEmpty(ws.Cells(r, "B").Value) Or Trim(ws.Cells(r, "B").Value) = "" Then
                errorMsg = errorMsg & "Procent Comision lipseste pentru ID Terminal: " & idTerm & vbCrLf
            End If
            ' Min and Max commission are optional - no validation needed
        End If
    Next r
    
    If errorMsg <> "" Then
        errorMsg = errorMsg & vbCrLf & "Completeaza datele lipsa in sheet-ul Comisioane."
        MsgBox errorMsg, vbCritical, "Eroare la validarea tabelului de comisioane"
        ValidateCommissionTable = False
    End If
End Function
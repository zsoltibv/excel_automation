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
            ' Check for missing values
            Dim completeMsg As String
            completeMsg = ". Completeaza datele lipsa in sheet-ul Comisioane." & vbCrLf

            If IsEmpty(ws.Cells(r, "B").Value) Or Trim(ws.Cells(r, "B").Value) = "" Then
                errorMsg = errorMsg & "Procent Comision lipseste pentru ID Terminal: " & idTerm & completeMsg
            End If

            If IsEmpty(ws.Cells(r, "C").Value) Or Trim(ws.Cells(r, "C").Value) = "" Then
                errorMsg = errorMsg & "Comision Minim lipseste pentru ID Terminal: " & idTerm & completeMsg
            End If

            If IsEmpty(ws.Cells(r, "D").Value) Or Trim(ws.Cells(r, "D").Value) = "" Then
                errorMsg = errorMsg & "Comision Maxim lipseste pentru ID Terminal: " & idTerm & completeMsg
            End If
        End If
    Next r
    
    If errorMsg <> "" Then
        MsgBox errorMsg, vbCritical, "Eroare la validarea tabelului de comisioane"
        ValidateCommissionTable = False
    End If
End Function
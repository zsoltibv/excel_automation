Option Explicit

'========================
' Write a single merged clsTxtFile into an Excel file
'========================
Public Sub WriteGroupedTxtFilesToExcel(txt As clsTxtFile, _
                                       outputPath As String, _
                                       ByVal startDate As Date, _
                                       ByVal endDate As Date)
    Dim wb As Workbook, ws As Worksheet
    Dim row As Long
    Dim tx As clsTransactionInfo
    
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    
    ' ===== HEADERS =====
    ws.Range("A1:M1").Value = Array( _
        "data_inreg", "data_op", _
        "valoare", "comision", _
        "nr_card", "retea", "tipc", _
        "cod_aut", "rrn", "document", _
        "id", "denumire", "cont")
    
    ' ===== COLUMN FORMATS =====
    ws.Columns(1).NumberFormat = "dd/mm/yyyy"
    ws.Columns(2).NumberFormat = "dd/mm/yyyy"
    ws.Columns(3).NumberFormat = "#,##0.00"
    ws.Columns(4).NumberFormat = "#,##0.00"
    ws.Columns(8).NumberFormat = "@"
    ws.Columns(9).NumberFormat = "@"
    ws.Columns(13).NumberFormat = "@"
    
    row = 2
    
    ' ===== DATA =====
    For Each tx In txt.Transactions
        ' ===== DATE FILTER =====
        If tx.DataOper >= startDate And tx.DataOper <= endDate Then
            ws.Cells(row, 1).Value = tx.DataInreg
            ws.Cells(row, 2).Value = tx.DataOper
            ws.Cells(row, 3).Value = tx.Valoare
            ws.Cells(row, 4).Value = tx.Comision
            ws.Cells(row, 5).Value = tx.NumarCard
            ws.Cells(row, 6).Value = tx.Retea
            ws.Cells(row, 7).Value = tx.TipC
            ws.Cells(row, 8).Value = tx.CodAut
            ws.Cells(row, 9).Value = tx.RRN
            ws.Cells(row, 10).Value = tx.Document
            ws.Cells(row, 11).Value = tx.IdTerm
            ws.Cells(row, 12).Value = tx.DenumireTerminal
            ws.Cells(row, 13).Value = tx.Cont
            row = row + 1
        End If
    Next tx
    
    ws.Columns.AutoFit
    wb.SaveAs outputPath, xlOpenXMLWorkbook
    wb.Close False
End Sub

'========================
' Export to excel each merged clsTxtFile in the grouped collection
'========================
Public Sub ExportGroupedFilesToExcel(grouped As Object, _
                                      outputFolder As String, _
                                      startDate As Date, _
                                      endDate As Date)
    Dim groupKey As Variant
    Dim outputPath As String
    Dim mergedTxt As clsTxtFile
    Dim datePart As String
    
    datePart = Format(startDate, "dd-mm-yyyy") & "_to_" & Format(endDate, "dd-mm-yyyy")
    
    For Each groupKey In grouped.Keys
        Set mergedTxt = grouped(groupKey)
        
        outputPath = outputFolder & "\" & _
                    CleanFileName(mergedTxt.Header.NumeComerciant) & "_" & _
                    PaymentTypeToString(mergedTxt.Header.Payment) & "_" & _
                    datePart & ".xlsx"
        
        Application.ScreenUpdating = False
        WriteGroupedTxtFilesToExcel mergedTxt, outputPath, startDate, endDate
        Application.ScreenUpdating = True
    Next groupKey
End Sub

Private Function CleanFileName(str As String) As String
    CleanFileName = Replace(str, " ", "_")
End Function
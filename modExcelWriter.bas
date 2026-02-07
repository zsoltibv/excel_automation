Option Explicit

'========================
' Write a collection of clsTxtFile into a single Excel file
'========================
Public Sub WriteGroupedTxtFilesToExcel(txtCollection As Collection, outputPath As String)
    Dim wb As Workbook, ws As Worksheet
    Dim row As Long
    Dim txt As clsTxtFile
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
    For Each txt In txtCollection
        For Each tx In txt.Transactions
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

            ws.Cells(row, 11).Value = txt.Header.IdTerm
            ws.Cells(row, 12).Value = txt.Header.DenumireTerminal
            ws.Cells(row, 13).Value = txt.Header.Cont

            row = row + 1
        Next tx
    Next txt

    ws.Columns.AutoFit
    wb.SaveAs outputPath, xlOpenXMLWorkbook
    wb.Close False
End Sub
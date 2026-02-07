Option Explicit

Public Sub WriteTxtFileToExcel(txt As clsTxtFile, outputPath As String)
    Dim wb As Workbook, ws As Worksheet
    Dim row As Long
    Dim tx As clsTransactionInfo

    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)

    ws.Range("A1:M1").Value = Array( _
    "data_inreg", "data_op", _
    "valoare", "comision", "nr_card", _
    "retea", "tipc", "cod_aut", _
    "rrn", "document", _
    "id", "denumire", "cont")

    row = 2

    For Each tx In txt.Transactions
        ws.Cells(row, 1).Value = tx.DataInreg
        ws.Cells(row, 2).Value = tx.DataOper
        ws.Cells(row, 3).Value = tx.Valoare
        ws.Cells(row, 4).Value = tx.Comision
        ws.Cells(row, 5).Value = tx.NumarCard
        ws.Cells(row, 6).Value = tx.Retea
        ws.Cells(row, 7).Value = tx.TipC
        ws.Cells(row, 8).Value = tx.CodAut
        ws.Cells(row, 9).NumberFormat = "@": ws.Cells(row, 9).Value = tx.RRN
        ws.Cells(row, 10).Value = tx.Document

        ws.Cells(row, 11).Value = txt.Header.IdTerm
        ws.Cells(row, 12).Value = txt.Header.DenumireTerminal
        ws.Cells(row, 13).Value = txt.Header.Cont

        row = row + 1
    Next tx

    ws.Columns.AutoFit
    wb.SaveAs outputPath, xlOpenXMLWorkbook
    wb.Close False
End Sub

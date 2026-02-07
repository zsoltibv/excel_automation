Attribute VB_Name = "Module1"
Option Explicit

Sub TxtToExcel()

    Dim fso As Object, folder As Object, file As Object
    Dim ts As Object
    Dim line As String
    Dim row As Long

    ' Hardcoded relative paths
    Dim inputFolder As String, outputFolder As String
    inputFolder = ThisWorkbook.Path & "\input"
    outputFolder = ThisWorkbook.Path & "\output"

    ' Create output folder if it doesn't exist
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outputFolder) Then fso.CreateFolder outputFolder

    Set folder = fso.GetFolder(inputFolder)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "txt" Then

            Dim wb As Workbook, ws As Worksheet
            Set wb = Workbooks.Add
            Set ws = wb.Sheets(1)

            ' Headers
            ws.Range("A1:M1").Value = Array( _
                "data_inreg", "data_op", _
                "valoare", "comision", "nr_card", _
                "retea", "tipc", "cod_aut", _
                "rrn", "document", _
                "id", "denumire", "cont")

            row = 2
            Set ts = file.OpenAsTextStream(1)

            ' Create header class instance
            Dim header As clsHeaderInfo
            Set header = New clsHeaderInfo

            ' Transaction class variable
            Dim tx As clsTransactionInfo

            Do While Not ts.AtEndOfStream
                line = ts.ReadLine

                ' Extract header info (once per file)
                If header.IdTerm = "" And InStr(line, "IdTerm:[") > 0 Then
                    header.IdTerm = Trim(Mid(line, InStr(line, "IdTerm:[") + 8))
                    header.IdTerm = Replace(header.IdTerm, "]", "")
                End If

                If header.DenumireTerminal = "" And Trim(line) Like "Denumire Terminal:*" Then
                    header.DenumireTerminal = Trim(Left(Mid(line, InStr(line, ":") + 1), 30))
                End If

                If header.Cont = "" And Trim(line) Like "Denumire Cont:*" Then
                    header.Cont = Trim(Mid(line, InStr(line, ":") + 1))
                End If

                ' Skip any line starting with "Referinta"
                If Trim(line) Like "Referinta:*" Then GoTo NextLine

                ' Transaction line starts with date
                If line Like "##/##/####*" Then
                    Set tx = New clsTransactionInfo
                    Set tx.Header = header    ' store header inside transaction
                    With tx
                        .DataInreg = Trim(Mid(line, 1, 10))
                        .DataOper = Trim(Mid(line, 12, 10))
                        .Valoare = Replace(Trim(Mid(line, 32, 14)), ",", "")
                        .Comision = Trim(Mid(line, 48, 12))
                        .NumarCard = Trim(Mid(line, 62, 18))
                        .Retea = Trim(Mid(line, 80, 5))
                        .TipC = Trim(Mid(line, 86, 5))
                        .CodAut = Trim(Mid(line, 95, 7))
                        .RRN = Trim(Mid(line, 102, 12))
                        .Document = Trim(Mid(line, 115))
                    End With

                    ' Write to worksheet
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

                    ' Add header info from transaction
                    ws.Cells(row, 11).Value = tx.Header.IdTerm
                    ws.Cells(row, 12).Value = tx.Header.DenumireTerminal
                    ws.Cells(row, 13).Value = tx.Header.Cont

                    row = row + 1
                End If

NextLine:
            Loop

            ts.Close
            ws.Columns.AutoFit

            ' Save workbook
            Dim outName As String
            outName = outputFolder & "\" & fso.GetBaseName(file.Name) & ".xlsx"
            wb.SaveAs outName, xlOpenXMLWorkbook
            wb.Close False

        End If
    Next file

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Import completed successfully", vbInformation

End Sub

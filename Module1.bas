Attribute VB_Name = "Module1"

Sub TxtToExcel()

    Dim fso As Object, folder As Object, file As Object
    Dim ts As Object
    Dim line As String
    Dim row As Long

    Dim inputFolder As String, outputFolder As String

    ' Pick INPUT folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select INPUT folder (TXT files)"
        If .Show <> -1 Then Exit Sub
        inputFolder = .SelectedItems(1)
    End With

    ' Pick OUTPUT folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select OUTPUT folder (Excel files)"
        If .Show <> -1 Then Exit Sub
        outputFolder = .SelectedItems(1)
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(inputFolder)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "txt" Then

            Dim wb As Workbook
            Dim ws As Worksheet
            Set wb = Workbooks.Add
            Set ws = wb.Sheets(1)

            ' Headers (add id, denumire, cont at the end)
            ws.Range("A1:M1").Value = Array( _
                "data_inreg", "data_op", _
                "valoare", "comision", "nr_card", _
                "retea", "tipc", "cod_aut", _
                "rrn", "document", _
                "id", "denumire", "cont")

            row = 2
            Set ts = file.OpenAsTextStream(1)

            ' Variables to store header info
            Dim idVal As String, denumireVal As String, contVal As String
            idVal = ""
            denumireVal = ""
            contVal = ""

            Do While Not ts.AtEndOfStream
                line = ts.ReadLine

                ' Extract header info (once per file)
                If idVal = "" And InStr(line, "IdTerm:[") > 0 Then
                    idVal = Trim(Mid(line, InStr(line, "IdTerm:[") + 8))
                    idVal = Replace(idVal, "]", "")
                End If

                If denumireVal = "" And Trim(line) Like "Denumire Terminal:*" Then
                    denumireVal = Trim(Left(Mid(line, InStr(line, ":") + 1), 30))
                End If

                If contVal = "" And Trim(line) Like "Denumire Cont:*" Then
                    contVal = Trim(Mid(line, InStr(line, ":") + 1))
                End If

                ' Skip any line starting with "Referinta"
                If Trim(line) Like "Referinta:*" Then GoTo NextLine

                ' Transaction line starts with date
                If line Like "##/##/####*" Then
                    ws.Cells(row, 1).Value = Trim(Mid(line, 1, 10))        ' DataInreg
                    ws.Cells(row, 2).Value = Trim(Mid(line, 12, 10))       ' DataOper
                    ws.Cells(row, 3).Value = Replace(Trim(Mid(line, 32, 14)), ",", "") ' SumaOper
                    ws.Cells(row, 4).Value = Trim(Mid(line, 48, 12))       ' Comision
                    ws.Cells(row, 5).Value = Trim(Mid(line, 62, 18))       ' NumarCard
                    ws.Cells(row, 6).Value = Trim(Mid(line, 80, 5))        ' Retea
                    ws.Cells(row, 7).Value = Trim(Mid(line, 86, 5))        ' TipC
                    ws.Cells(row, 8).Value = Trim(Mid(line, 95, 7))        ' CodAut
                    ws.Cells(row, 9).NumberFormat = "@": ws.Cells(row, 9).Value = Trim(Mid(line, 102, 12)) ' RRN
                    ws.Cells(row, 10).Value = Trim(Mid(line, 115))         ' Document

                    ' Add the 3 new columns
                    ws.Cells(row, 11).Value = idVal
                    ws.Cells(row, 12).Value = denumireVal
                    ws.Cells(row, 13).Value = contVal

                    row = row + 1
                End If

NextLine:
            Loop

            ts.Close
            ws.Columns.AutoFit

            ' Save with same name as TXT
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

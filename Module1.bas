Attribute VB_Name = "Module1"
Option Explicit

Sub TxtToExcel()
    Dim inputFolder As String, outputFolder As String
    Dim fso As Object, folder As Object, file As Object
    Dim transactions As Collection

    inputFolder = ThisWorkbook.Path & "\input"
    outputFolder = ThisWorkbook.Path & "\output"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outputFolder) Then fso.CreateFolder outputFolder

        Set folder = fso.GetFolder(inputFolder)

        Application.ScreenUpdating = False
        Application.DisplayAlerts = False

        For Each file In folder.Files
            If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
                Set transactions = ParseTxtFile(file)
                WriteTransactionsToExcel transactions, outputFolder & "\" & fso.GetBaseName(file.Name) & ".xlsx"
            End If
        Next file

        Application.ScreenUpdating = True
        Application.DisplayAlerts = True

        MsgBox "Import completed successfully", vbInformation
End Sub
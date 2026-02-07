Attribute VB_Name = "Module1"
Option Explicit

Public Sub TxtToExcelGroupedFiles()
    Dim inputFolder As String, outputFolder As String
    Dim fso As Object, folder As Object, file As Object
    Dim txt As clsTxtFile
    Dim txtList As New Collection
    Dim grouped As Object 
    Dim key As String

    inputFolder = ThisWorkbook.Path & "\input"
    outputFolder = ThisWorkbook.Path & "\output"

    ' ===== Create output folder if missing =====
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outputFolder) Then fso.CreateFolder outputFolder
    Set folder = fso.GetFolder(inputFolder)

    ' ===== Parse all TXT files =====
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
            Set txt = ParseTxtFile(file)
            txtList.Add txt
        End If
    Next file

    ' ===== Group TXT files by IdComer + Payment =====
    Set grouped = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To txtList.Count
        Set txt = txtList(i)
        key = txt.Header.IdComer & "_" & txt.Header.Payment ' Unique key per group
        If Not grouped.Exists(key) Then
            grouped.Add key, New Collection
        End If
        grouped(key).Add txt
    Next i

    ' ===== Write each group to a separate Excel file =====
    Dim groupKey As Variant
    Dim outputPath As String
    Dim firstTxt As clsTxtFile
    Dim paymentText As String

    For Each groupKey In grouped.Keys
        Set firstTxt = grouped(groupKey)(1)
        
        ' Determine payment type text
        Select Case firstTxt.Header.Payment
            Case PaymentType.POS
                paymentText = "POS"
            Case PaymentType.ECOMMERCE
                paymentText = "ECOMMERCE"
            Case Else
                paymentText = "UNKNOWN"
        End Select
        
        ' Build output filename
        outputPath = outputFolder & "\" & CleanFileName(firstTxt.Header.NumeComerciant) & "_" & paymentText & ".xlsx"
        
        ' Write group to Excel
        Application.ScreenUpdating = False
        WriteGroupedTxtFilesToExcel grouped(groupKey), outputPath
        Application.ScreenUpdating = True
    Next groupKey

    MsgBox "All TXT files grouped and written to separate Excel files.", vbInformation
End Sub

Private Function CleanFileName(str As String) As String
    Dim invalidChars As Variant
    Dim ch As Variant
    invalidChars = Array(" ")
    CleanFileName = str
    For Each ch In invalidChars
        CleanFileName = Replace(CleanFileName, ch, "_")
    Next ch
End Function
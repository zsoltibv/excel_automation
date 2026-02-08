Attribute VB_Name = "Module1"
Option Explicit

Public Sub RunTxtImportWithDateFilter()
    Dim frm As New frmDateFilter
    
    frm.IsCancelled = True
    frm.Show vbModal
    
    If frm.IsCancelled Then Exit Sub
    
    TxtToExcelGroupedFiles frm.StartDate, frm.EndDate
End Sub

Public Sub TxtToExcelGroupedFiles(ByVal startDate As Date, ByVal endDate As Date)
    Dim inputFolder As String, outputFolder As String
    Dim fso As Object, folder As Object, file As Object
    Dim txt As clsTxtFile
    Dim txtList As New Collection
    Dim grouped As Object
    Dim commissions As Object
    
    inputFolder = ThisWorkbook.Path & "\input"
    outputFolder = ThisWorkbook.Path & "\output"
    
    ' ===== Create output folder if missing =====
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outputFolder) Then fso.CreateFolder outputFolder
    Set folder = fso.GetFolder(inputFolder)
    
    ' ===== Parse commission table FIRST =====
    Set commissions = ParseCommissionTable()
    If commissions Is Nothing Then Exit Sub 
    
    ' ===== Parse all TXT files =====
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
            Set txt = ParseTxtFile(file)
            txtList.Add txt
        End If
    Next file
    
    ' ===== Group TXT files with filtering and commission calculation =====
    Set grouped = GroupTxtFiles(txtList, startDate, endDate, commissions)
    If grouped Is Nothing Then Exit Sub 
    
    ' ===== Export grouped files to Excel =====
    ExportGroupedFilesToExcel grouped, outputFolder, startDate, endDate
    
    MsgBox "Operatiunea s-a finalizat cu succes." & vbCrLf & _
       "Fisierele TXT au fost grupate si exportate in fisiere Excel separate.", _
       vbInformation, "Finalizat"
End Sub
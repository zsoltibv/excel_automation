Attribute VB_Name = "modBackup"
Option Explicit

Public Sub BackupAndCleanInputFolder(inputFolder As String, startDate As Date, endDate As Date)
    Dim fso As Object
    Dim backupFolder As String
    Dim datePart As String
    Dim file As Object
    Dim folder As Object
    Dim counter As Long
    Dim baseName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(inputFolder)
    
    ' Create backup folder name with date
    datePart = Format(startDate, "dd-mm-yyyy") & "_to_" & Format(endDate, "dd-mm-yyyy")
    baseName = ThisWorkbook.Path & "\backup_" & datePart
    backupFolder = baseName
    
    ' Handle folder name collision by adding counter
    counter = 1
    Do While fso.FolderExists(backupFolder)
        backupFolder = baseName & "_" & counter
        counter = counter + 1
    Loop
    
    ' Create backup folder
    fso.CreateFolder backupFolder
    
    ' Copy all TXT files to backup folder
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
            fso.CopyFile file.Path, backupFolder & "\" & file.Name, True
        End If
    Next file
    
    ' Delete all TXT files from input folder
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
            fso.DeleteFile file.Path, True
        End If
    Next file
End Sub
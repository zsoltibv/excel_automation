Attribute VB_Name = "modGrouping"

Option Explicit

Public Function GroupTxtFiles(txtList As Collection, _
                              startDate As Date, _
                              endDate As Date, _
                              commissions As Object) As Object
    Dim grouped As Object
    Dim txt As clsTxtFile
    Dim mergedTxt As clsTxtFile
    Dim key As String
    Dim i As Long
    
    Set grouped = CreateObject("Scripting.Dictionary")
    
    For i = 1 To txtList.Count
        Set txt = txtList(i)
        key = txt.Header.IdComer & "_" & txt.Header.Payment
        
        If Not grouped.Exists(key) Then
            Set mergedTxt = New clsTxtFile
            ' Copy header info from first file
            mergedTxt.Header.NumeComerciant = txt.Header.NumeComerciant
            mergedTxt.Header.IdComer = txt.Header.IdComer
            mergedTxt.Header.Payment = txt.Header.Payment
            grouped.Add key, mergedTxt
        End If
        
        grouped(key).MergeTxtFileFiltered txt, startDate, endDate, commissions
    Next i
    
    Set GroupTxtFiles = grouped
End Function
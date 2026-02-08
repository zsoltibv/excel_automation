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
    Dim hasTransactions As Boolean
    
    Set grouped = CreateObject("Scripting.Dictionary")
    
    For i = 1 To txtList.Count
        Set txt = txtList(i)
        key = txt.Header.IdComer & "_" & txt.Header.Payment

        ' Validate IdTerm exists in commission table
        If Not ValidateIdTerm(txt.Header.IdTerm, commissions) Then
            Set GroupTxtFiles = Nothing
            Exit Function
        End If
        
        If Not grouped.Exists(key) Then
            Set mergedTxt = New clsTxtFile
            ' Copy header info from first file
            mergedTxt.Header.NumeComerciant = txt.Header.NumeComerciant
            mergedTxt.Header.IdComer = txt.Header.IdComer
            mergedTxt.Header.Payment = txt.Header.Payment
            grouped.Add key, mergedTxt
        End If
        
        grouped(key).MergeTxtFileFiltered txt, startDate, endDate, commissions

        ' Check if any transactions were added
        If grouped(key).Transactions.Count > 0 Then
            hasTransactions = True
        End If
    Next i

    ' Check if any transactions exist
    If Not hasTransactions Then
        MsgBox "Nu au fost gasite tranzactii in intervalul selectat sau folderul input este gol." & vbCrLf & _
            "Verifica datele si incearca din nou.", _
            vbExclamation, "Nicio tranzactie gasita"
        Set GroupTxtFiles = Nothing
        Exit Function
    End If
    
    Set GroupTxtFiles = grouped
End Function

Private Function ValidateIdTerm(idTerm As String, commissions As Object) As Boolean
    ValidateIdTerm = True
    
    If Not commissions.Exists(idTerm) Then
        MsgBox "ID Terminal '" & idTerm & "' nu exista in sheet-ul Comisioane." & vbCrLf & _
               "Adauga acest terminal in sheet-ul Comisioane.", _
               vbCritical, "Eroare validare ID Terminal"
        ValidateIdTerm = False
    End If
End Function
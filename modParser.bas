Option Explicit

Public Function ParseTxtFile(file As Object) As clsTxtFile
    Dim ts As Object, line As String
    Dim txt As clsTxtFile
    Dim tx As clsTransactionInfo

    Set txt = New clsTxtFile
    txt.FileName = file.Name

    Set ts = file.OpenAsTextStream(1)

    Do While Not ts.AtEndOfStream
        line = ts.ReadLine

        ' ===== HEADER =====
        With txt.Header
            If .IdTerm = "" And InStr(line, "IdTerm:[") > 0 Then
                .IdTerm = Replace(Trim(Mid(line, InStr(line, "IdTerm:[") + 8)), "]", "")
            End If

            If .DenumireTerminal = "" And Trim(line) Like "Denumire Terminal:*" Then
                .DenumireTerminal = Trim(Left(Mid(line, InStr(line, ":") + 1), 30))
            End If

            If .Cont = "" And Trim(line) Like "Denumire Cont:*" Then
                .Cont = Trim(Mid(line, InStr(line, ":") + 1))
            End If
        End With

        If Trim(line) Like "Referinta:*" Then Goto NextLine

            ' ===== TRANSACTION =====
            If line Like "##/##/####*" Then
                Set tx = New clsTransactionInfo
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

                txt.AddTransaction tx
            End If

 NextLine:
        Loop

        ts.Close
        Set ParseTxtFile = txt
End Function

' modParser
Option Explicit

Public Function ParseTxtFile(file As Object) As Collection
    Dim ts As Object, line As String
    Dim header As clsHeaderInfo
    Dim tx As clsTransactionInfo
    Dim transactions As Collection

    Set ts = file.OpenAsTextStream(1)
    Set transactions = New Collection
    Set header = New clsHeaderInfo

    Do While Not ts.AtEndOfStream
        line = ts.ReadLine

        ' Extract header info
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

        ' Skip lines starting With "Referinta"
        If Trim(line) Like "Referinta:*" Then Goto NextLine

            ' Transaction lines start With date
            If line Like "##/##/####*" Then
                Set tx = New clsTransactionInfo
                Set tx.Header = header
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
                transactions.Add tx
            End If

 NextLine:
        Loop

        ts.Close
        Set ParseTxtFile = transactions
End Function

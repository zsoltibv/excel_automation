Option Explicit

Public Function ParseTxtFile(file As Object) As clsTxtFile
    Dim ts As Object
    Dim txt As clsTxtFile
    
    Set txt = New clsTxtFile
    txt.FileName = file.Name
    
    Set ts = file.OpenAsTextStream(1)
    
    ' ===== Parse HEADER =====
    ParseHeader ts, txt
    
    ' ===== Parse TRANSACTIONS =====
    ParseTransactions ts, txt
    
    ts.Close
    Set ParseTxtFile = txt
End Function

'========================
' Parse header fields until all are filled
'========================
Private Sub ParseHeader(ts As Object, txt As clsTxtFile)
    Dim line As String
    
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        
        With txt.Header
            If .IdTerm = "" And InStr(line, "IdTerm:[") > 0 Then
                .IdTerm = Trim(Split(Split(line, "IdTerm:[")(1), " ")(0))
                .IdTerm = Trim(Split(.IdTerm, "]")(0))

                ' ===== Set Payment Type based on first character =====
                Select Case Left(.IdTerm, 1)
                    Case "5"
                        .Payment = PaymentType.ECOMMERCE
                    Case "6"
                        .Payment = PaymentType.POS
                    Case Else
                        .Payment = PaymentType.UNKNOWN
                End Select
            End If

            If .IdComer = "" And InStr(line, "IdComer:[") > 0 Then
                .IdComer = Trim(Split(Split(line, "IdComer:[")(1), "]")(0))
            End If

            If .DenumireTerminal = "" And Trim(line) Like "Denumire Terminal:*" Then
                .DenumireTerminal = RTrim(Left(Mid(line, InStr(line, ":") + 1), 35))
            End If

            If .NumeComerciant = "" And Trim(line) Like "Nume Comerciant:*" Then
                .NumeComerciant = RTrim(Left(Mid(line, InStr(line, ":") + 1), 34))
            End If

            If .Cont = "" And InStr(line, "Cont:[") > 0 Then
                .Cont = Trim(Split(Split(line, "Cont:[")(1), "]")(0))
            End If
        End With
        
        ' Stop once all header fields are filled
        If txt.Header.IdTerm <> "" And _
           txt.Header.IdComer <> "" And _
           txt.Header.DenumireTerminal <> "" And _
           txt.Header.NumeComerciant <> "" And _
           txt.Header.Cont <> "" Then Exit Do
    Loop
End Sub

'========================
' Parse transaction lines
'========================
Private Sub ParseTransactions(ts As Object, txt As clsTxtFile)
    Dim line As String
    Dim tx As clsTransactionInfo
    
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        If Trim(line) Like "Referinta:*" Then GoTo NextLine
        
        If line Like "##/##/####*" Then
            Set tx = New clsTransactionInfo
            With tx
                .DataInreg = ParseDateDMY(Mid(line, 1, 10))
                .DataOper  = ParseDateDMY(Mid(line, 12, 10))
                .Valoare   = CCur(Replace(Mid(line, 32, 14), ",", ""))
                .Comision  = CCur(Replace(Mid(line, 48, 12), ",", ""))
                .NumarCard = Trim(Mid(line, 62, 18))
                .Retea     = Trim(Mid(line, 80, 5))
                .TipC      = Trim(Mid(line, 86, 5))
                .CodAut    = Trim(Mid(line, 95, 7))
                .RRN       = Trim(Mid(line, 102, 12))
                .Document  = Trim(Mid(line, 115))

                ' ===== Save header info with transaction =====
                .IdTerm = txt.Header.IdTerm
                .DenumireTerminal = txt.Header.DenumireTerminal
                .Cont = txt.Header.Cont
            End With
            txt.AddTransaction tx
        End If
        
NextLine:
    Loop
End Sub

'========================
' Helper to parse dd/mm/yyyy into Date
'========================
Public Function ParseDateDMY(ByVal value As String) As Date
    ParseDateDMY = DateSerial( _
        CInt(Mid(value, 7, 4)), _
        CInt(Mid(value, 4, 2)), _
        CInt(Mid(value, 1, 2)) _
    )
End Function
Attribute VB_Name = "modValidation"

Option Explicit

Public Function TryParseDateDMY(ByVal s As String, ByRef outDate As Date, ByRef errMsg As String) As Boolean
    Dim d As Integer, m As Integer, y As Integer
    Dim tmp As Date

    s = Trim(s)

    ' Format: dd/mm/yyyy
    If Not s Like "##/##/####" Then
        errMsg = "Format invalid. Foloseste formatul: zz/ll/aaaa."
        Exit Function
    End If

    ' Numeric parts
    On Error GoTo InvalidNumber
    d = CInt(Mid(s, 1, 2))
    m = CInt(Mid(s, 4, 2))
    y = CInt(Mid(s, 7, 4))
    On Error GoTo 0

    If m < 1 Or m > 12 Then
        errMsg = "Luna trebuie sa fie intre 01 si 12."
        Exit Function
    End If

    If d < 1 Or d > 31 Then
        errMsg = "Ziua trebuie sa fie intre 01 si 31."
        Exit Function
    End If

    ' Real calendar validation (kills 31/02)
    tmp = DateSerial(y, m, d)

    If Day(tmp) <> d Or Month(tmp) <> m Or Year(tmp) <> y Then
        errMsg = "Data introdusa nu exista in calendar."
        Exit Function
    End If

    outDate = tmp
    TryParseDateDMY = True
    Exit Function

InvalidNumber:
    errMsg = "Data contine valori invalide."
End Function
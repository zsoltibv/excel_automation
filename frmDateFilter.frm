VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDateFilter 
   Caption         =   "UserForm1"
   ClientHeight    =   1935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6165
   OleObjectBlob   =   "frmDateFilter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDateFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' UserForm to get date range input from user
Public StartDate As Date
Public EndDate As Date
Public IsCancelled As Boolean

Private Sub btnOK_Click()
    Dim msg As String

    If Not TryParseDateDMY(txtStartDate.Value, StartDate, msg) Then
        MsgBox "Data de inceput invalida:" & vbCrLf & msg, vbExclamation
        txtStartDate.SetFocus
        Exit Sub
    End If

    If Not TryParseDateDMY(txtEndDate.Value, EndDate, msg) Then
        MsgBox "Data de sfarsit invalida:" & vbCrLf & msg, vbExclamation
        txtEndDate.SetFocus
        Exit Sub
    End If

    If EndDate < StartDate Then
        MsgBox "Data de sfarsit trebuie sa fie dupa data de inceput.", vbExclamation
        txtEndDate.SetFocus
        Exit Sub
    End If

    ' Save commision settings globally
    UseCommission = commCheckbox.Value

    IsCancelled = False
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    txtStartDate.Value = Format(DateSerial(Year(Date), Month(Date), 1), "dd/mm/yyyy")
    txtEndDate.Value = Format(Date, "dd/mm/yyyy")
End Sub
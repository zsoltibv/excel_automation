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
    On Error GoTo InvalidDate
    
    StartDate = ParseDateDMY(txtStartDate.Value)
    EndDate = ParseDateDMY(txtEndDate.Value)
    
    If EndDate < StartDate Then
        MsgBox "End date must be after Start date.", vbExclamation
        Exit Sub
    End If
    
    IsCancelled = False
    Me.Hide
    Exit Sub

InvalidDate:
    MsgBox "Please enter valid dates (dd/mm/yyyy).", vbExclamation
End Sub

Private Sub UserForm_Initialize()
    txtStartDate.Value = Format(DateSerial(Year(Date), Month(Date), 1), "dd/mm/yyyy")
    txtEndDate.Value = Format(Date, "dd/mm/yyyy")
End Sub
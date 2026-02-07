VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTxtFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mHeader As clsHeaderInfo
Private mTransactions As Collection
Private mFileName As String

Private Sub Class_Initialize()
    Set mHeader = New clsHeaderInfo
    Set mTransactions = New Collection
End Sub

' =====================
' Properties
' =====================
Public Property Get Header() As clsHeaderInfo
    Set Header = mHeader
End Property

Public Property Get Transactions() As Collection
    Set Transactions = mTransactions
End Property

Public Property Get FileName() As String
    FileName = mFileName
End Property

Public Property Let FileName(ByVal value As String)
    mFileName = value
End Property

' =====================
' Methods
' =====================
Public Sub AddTransaction(tx As clsTransactionInfo)
    mTransactions.Add tx
End Sub


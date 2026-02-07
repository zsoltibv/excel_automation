VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHeaderInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clsHeaderInfo
Option Explicit

' Header fields
Public IdTerm As String
Public IdComer As String
Public DenumireTerminal As String
Public NumeComerciant As String
Public Cont As String
Public Payment As PaymentType

' ===== Payment Type Enum =====
Public Enum PaymentType
    POS = 1
    ECOMMERCE = 2
    UNKNOWN = 3
End Enum
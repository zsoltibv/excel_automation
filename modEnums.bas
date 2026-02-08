Attribute VB_Name = "modEnums"
Option Explicit

Public Enum PaymentType
    POS = 1
    ECOMMERCE = 2
    UNKNOWN = 3
End Enum

Public Function PaymentTypeToString(pt As PaymentType) As String
    Select Case pt
        Case PaymentType.POS
            PaymentTypeToString = "POS"
        Case PaymentType.ECOMMERCE
            PaymentTypeToString = "ECOMMERCE"
        Case Else
            PaymentTypeToString = "UNKNOWN"
    End Select
End Function
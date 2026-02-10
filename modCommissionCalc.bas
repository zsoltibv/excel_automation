Attribute VB_Name = "modCommissionCalc"

Public Function CalculateCommission( _
    ByVal value As Currency, _
    ByVal comm As clsCommission) As Currency
    
    Dim result As Currency
    result = (value * comm.CommissionPercent) / 100@
    
    If comm.MaxCommission > 0@ And result > comm.MaxCommission Then
        result = comm.MaxCommission
    ElseIf comm.MinCommission > 0@ And result < comm.MinCommission Then
        result = comm.MinCommission
    End If
    
    CalculateCommission = CCur(Application.WorksheetFunction.Round(result, 2))
End Function
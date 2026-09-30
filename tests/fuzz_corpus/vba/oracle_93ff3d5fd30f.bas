Public Function InvoiceTotal(ByVal Subtotal As Currency, Optional ByVal TaxRate As Double = 0.08) As Currency
    InvoiceTotal = Subtotal + (Subtotal * TaxRate)
End Function

Public Sub XlideOracleEntry()
    Dim total As Double
    Dim total2 As Double
    total = 100
    total2 = InvoiceTotal(total, )
End Sub

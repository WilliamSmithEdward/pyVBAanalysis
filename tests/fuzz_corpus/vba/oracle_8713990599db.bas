Public Function InvoiceTotal(ByVal Subtotal As Currency, ByVal TaxRate As Double) As Currency
    InvoiceTotal = Subtotal + (Subtotal * TaxRate)
End Function

Public Sub XlideOracleEntry()
    Dim total As Double
    total = InvoiceTotal 100, 0.08
End Sub

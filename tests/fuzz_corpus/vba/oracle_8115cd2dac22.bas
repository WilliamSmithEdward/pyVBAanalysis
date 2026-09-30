Private Enum RuntimeArgStart
    BadStart = 0
End Enum

Public Sub XlideOracleEntry()
    Dim value As String
    value = Mid$("abcdef", BadStart, 1)
End Sub

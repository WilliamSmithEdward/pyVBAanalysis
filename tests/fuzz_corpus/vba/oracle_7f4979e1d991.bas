Public Sub Mutate(ByRef value As Long)
    value = value + 1
End Sub

Public Sub XlideOracleEntry()
    Dim amount As Long
    Mutate amount
End Sub

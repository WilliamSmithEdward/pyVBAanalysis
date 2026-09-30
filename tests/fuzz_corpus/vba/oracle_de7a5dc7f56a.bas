Public Sub XlideOracleEntry()
    Dim n As Long
    n = 1
    On n GoTo 100, 200
    Exit Sub
100
    Debug.Print "first"
    Exit Sub
200
    Debug.Print "second"
End Sub

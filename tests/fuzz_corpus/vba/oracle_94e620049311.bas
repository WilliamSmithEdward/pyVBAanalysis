Option Explicit

Friend Sub InternalOnly()
    Debug.Print "friend"
End Sub

Public Sub XlideOracleEntry()
    InternalOnly
End Sub

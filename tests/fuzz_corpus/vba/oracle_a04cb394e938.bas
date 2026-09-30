Option Explicit
Private Sub CloseAll()
    Close
End Sub

Function Main() As Long
    Dim p As String
    p = Environ$("TEMP") & "\xlide_files.txt"
    Open p For Output As #1
    Reset
    Open p For Output As #1
    CloseAll
    Open p For Output As #1
    Close #1
    Open p For Output As #&H1
    Open p & "2" For Output As #&H2
    Close #&H1, #&H2
    Main = Logged(p)
End Function

Private Function Logged(ByVal p As String) As Long
    On Error GoTo Failed
    Open p For Output As #1
    Print #1, "start"
    Err.Raise 1000, "Logged", "stop"
    Close #1
    Logged = 1
    Exit Function
Failed:
    Print #1, "failed: " & Err.Description
    Close #1
    Logged = 2
End Function

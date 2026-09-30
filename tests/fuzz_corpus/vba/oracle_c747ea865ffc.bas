Option Explicit
Function Main() As Long
    Dim p As String
    p = Environ$("TEMP") & "\xlide_files_b.txt"
    Open p For Output As #1
    Reset
    Print #1, "x"
    Main = 1
End Function

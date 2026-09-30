Option Explicit
Function Main() As Variant
    Dim f As Integer
    f = FreeFile
    Open Environ$("TEMP") & "\xlide_123c.txt" For Output As #f
    Close #f
    Print #f, "late"
    Main = 1
End Function

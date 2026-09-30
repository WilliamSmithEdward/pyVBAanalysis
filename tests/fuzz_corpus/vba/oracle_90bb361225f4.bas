Option Explicit
Function Main() As Variant
    Dim f As Integer, s As String
    f = FreeFile
    Open Environ$("TEMP") & "\xlide_123e.txt" For Output As #f
    Line Input #f, s
    Close #f
    Main = 1
End Function

Option Explicit
Function Main() As Long
    Dim b As Byte, i As Integer
    For b = 0 To 255
    Next b
    For i = 1 To 32767
        i = i
    Next i
    Main = 65535
End Function

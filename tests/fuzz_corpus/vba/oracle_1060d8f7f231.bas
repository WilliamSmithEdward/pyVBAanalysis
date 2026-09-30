Option Explicit
Private Type TBuf
    n As Long
    b() As Byte
End Type
Public Sub T()
    Dim buf As TBuf
    ReDim buf.b(0 To 3)
End Sub

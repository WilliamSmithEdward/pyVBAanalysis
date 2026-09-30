Option Explicit
Private mSize As Long
Public Property Get Size() As Long
    Size = mSize
End Property
Public Property Let Size(ByVal v As Integer)
    mSize = v
End Property

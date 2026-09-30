Option Explicit
Private mValue As Variant
Public Property Get Item(ByVal Index As Variant) As Variant
    If IsObject(mValue) Then
        Set Item = mValue
    Else
        Item = mValue
    End If
End Property
Public Property Let Item(ByVal Index As Variant, ByVal Value As Variant)
    mValue = Value
End Property
Public Property Set Item(ByVal Index As Variant, ByVal Value As Object)
    Set mValue = Value
End Property

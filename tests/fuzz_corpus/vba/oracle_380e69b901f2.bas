Option Explicit
Function Main() As Variant
    Dim n As Long
    If n = 0 Then
        If n = 0 Then
            Main = "inner"
#If VBA7 Then
        End If
#Else
        End If
#End If
    Else
        Main = "outer else"
    End If
End Function

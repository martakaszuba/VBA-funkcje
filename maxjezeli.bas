Attribute VB_Name = "Module1"
Function MAXJEZELI(str As String)
num = 0
For Each el In Range("B2:C20")
If el.Value = str Then
If el.Offset(0, 1).Value > num Then
num = el.Offset(0, 1)
End If
End If
Next
MAXJEZELI = num
End Function

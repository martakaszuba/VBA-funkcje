Attribute VB_Name = "Module1"
Function maxjezeli(rng As Range, txt As String)
Dim num As Integer
num = 0
For Each el In rng
If el.Value = txt Then
If el.Offset(0, 1).Value > num Then
num = el.Offset(0, 1).Value
End If
End If
Next
maxjezeli = num
End Function

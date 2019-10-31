Attribute VB_Name = "Module1"
Function JAKWIELE(rng As Range, txt As String)
  
  '             A                      B
  ' 1  Komedia, Dramat, Romans      = JAKWIELE(A1:A2;"Komedia")
  ' 2  Komedia, Romans
  
Dim str As String
str = ""
Dim num As Integer
num = 0
For Each el In rng
Lstr = Split(el.Value, ", ")
For i = LBound(Lstr) To UBound(Lstr)
If Lstr(i) = txt Then
num = num + 1
End If
Next
Next
JAKWIELE = num
End Function


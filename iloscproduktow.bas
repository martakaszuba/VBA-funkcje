Attribute VB_Name = "Module1"
Function ILOSCPRODUKTOW(str As String)
  '            A                          B
  ' 1     jajko,maslo,pomidor       =ILOSCPRODUKTOW(A1) => 3

Dim num As Integer
num = 0
g = Len(str)
For i = 1 To g
If Mid(str, i, 1) = "," Then
num = num + 1
End If
Next
ILOSCPRODUKTOW = num + 1
End Function

Attribute VB_Name = "Module1"
Function OSTATNI(str As String, del As String)
arr = Split(str, del)
OSTATNI = arr(UBound(arr))
End Function

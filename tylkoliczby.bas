Attribute VB_Name = "Module1"
Function TYLKOLICZBY(str As String)

'         A               B
' 1     wojtek34       =TYLKOLICZBY(A1) =>34

Dim txt As String
txt = ""
For i = 1 To Len(str)
If IsNumeric(Mid(str, i, 1)) = True Then
txt = txt & Mid(str, i, 1)
End If
Next
TYLKOLICZBY = txt
End Function

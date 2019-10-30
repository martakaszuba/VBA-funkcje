Attribute VB_Name = "Module1"
Function POLACZONYSTR(str As String)
Dim finalStr As String
finalStr = ""
For Each el In Range("A2:A14")
If el.Value = str Then
finalStr = finalStr & el.Offset(0, 1).Value & ", "
End If
Next
POLACZONYSTR = Mid(finalStr, 1, Len(finalStr) - 2)
End Function

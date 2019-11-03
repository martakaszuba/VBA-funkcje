Attribute VB_Name = "Module1"
Function ILEUNIKATOWYCH(rng As Range)

'        A            B
' 1   Sygnity      =ILEUNIKATOWYCH(A1:A4) => 3
' 2   Google
' 3   Comarch
' 4   Google

Dim myCol As Collection
Set myCol = New Collection
Dim num As Integer
Dim bool As Boolean
num = 0
For Each el In rng
bool = True
If num = 0 Then
myCol.Add el.Value
Else
For Each Item In myCol
If Item = el.Value Then
bool = False
Exit For
End If
Next
If bool Then
myCol.Add el.Value
End If
End If
num = num + 1
Next
ILEUNIKATOWYCH = myCol.Count
End Function

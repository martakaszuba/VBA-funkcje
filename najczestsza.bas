Attribute VB_Name = "Module1"
Function NAJCZESTSZA(rng As Variant)
  'oblicza najczęściej występującą wartość/wartości

'                A                     B
' 1           Wojtek              =NAJCZESTSZA(A1:A5) =>Wojtek,Magda
' 2           Magda
' 3           Wojtek
' 4           Rafał
' 5           Magda

Dim num As Integer
num = 1
Dim str As String
str = ""
Dim count As Integer
count = 0
Dim bool As Boolean
Dim finalStr As String
finalStr = ""
Dim c As Variant
Dim maxc As Integer
maxc = 1
Dim myCol As Collection
Set myCol = New Collection

For Each el In rng
bool = True
c = Application.WorksheetFunction.CountIf(rng, el.Value)
For Each Item In myCol
If count = 0 Then
myCol.Add el.Value & ":" & c
Else
If el.Value & ":" & c = Item Then
bool = False
End If
End If
Next
If bool Then
myCol.Add el.Value & ":" & c
End If
count = count + 1
Next

For Each el2 In myCol
f = Split(el2, ":")
If f(1) > maxc Then
maxc = f(1)
End If
Next

If maxc = 1 Then
NAJCZESTSZA = "Same unikaty"
Else
For Each el3 In myCol
g = Split(el3, ":")
If g(1) = maxc Then
finalStr = finalStr & "," & g(0)
End If
Next
NAJCZESTSZA = Mid(finalStr, 2, 500)
End If
End Function

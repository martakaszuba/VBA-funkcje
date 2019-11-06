Attribute VB_Name = "Module1"
Function MAXJEZELI(zakres As Range, kryteria As String, odleglosc As Integer)

'zakres - kolumna tekstowa
'kryteria - szukany tekst
'odleglosc - odleglosc miedzy kolumną z liczbami a kolumną tekstową

'      A              B                  C
'1    Anna           3400           =MAXJEZELI(A1:A5;"Anna";1) => 5700
'2    Wojtek         4500
'3    Anna           5700
'4    Rafał          4300
'5    Anna           2800

Dim num As Integer
num = 0
For Each el In zakres
If el.Value = kryteria Then
If el.Offset(0, odleglosc).Value > num Then
num = el.Offset(0, odleglosc).Value
End If
End If
Next
MAXJEZELI = num
End Function

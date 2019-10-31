Attribute VB_Name = "Module1"
Function JAKIDZIEN(str As String)
daynum = Weekday(str, 2)
If daynum = 1 Then
daystr = "poniedzia³ek"
ElseIf daynum = 2 Then
daystr = "wtorek"
ElseIf daynum = 3 Then
daystr = "œroda"
ElseIf daynum = 4 Then
daystr = "czwartek"
ElseIf daynum = 5 Then
daystr = "pi¹tek"
ElseIf daynum = 6 Then
daystr = "sobota"
Else
daystr = "niedziela"
End If
JAKIDZIEN = daystr
End Function

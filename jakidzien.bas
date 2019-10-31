Attribute VB_Name = "Module1"
Function JAKIDZIEN(str As String)
daynum = Weekday(str, 2)
If daynum = 1 Then
daystr = "poniedziałek"
ElseIf daynum = 2 Then
daystr = "wtorek"
ElseIf daynum = 3 Then
daystr = "środa"
ElseIf daynum = 4 Then
daystr = "czwartek"
ElseIf daynum = 5 Then
daystr = "piątek"
ElseIf daynum = 6 Then
daystr = "sobota"
Else
daystr = "niedziela"
End If
JAKIDZIEN = daystr
End Function

Attribute VB_Name = "Module1"
Sub multiply()
Range("D5").Select
Do Until ActiveCell.Value = ""
ActiveCell.Offset(0, 1).Value = ActiveCell.Value * 2
ActiveCell.Offset(1, 0).Select
Loop

End Sub

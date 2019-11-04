Attribute VB_Name = "Module1"
Function OSTATNI(str As String, del As String)
  'oblicza ostatni element w danym tekście
  
  '                   A                                     B
  '1   C:\Maciej\fold\Mazowieckie\Warszawa           =OSTATNI(A1;"\") => Warszawa
  
arr = Split(str, del)
OSTATNI = arr(UBound(arr))
End Function

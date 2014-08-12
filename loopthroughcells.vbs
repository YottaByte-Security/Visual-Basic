Private Sub CommandButton1_Click()
Dim Day As Integer
Dim Count As Integer
Dim Index As Integer
For Day = 3 To 19
  For Count = 1 To 100
    Index = Day * 100 + Count
    Range("A" & Index).Value = Day
   Next Count
Next Day
End Sub

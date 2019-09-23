```VB
' Vloži riadok za každý ďalší riadok do výberu
' Insert A Row After Every Other Row in the Selection
Sub InsertAlternateRows()
	Dim rng As Range
	Dim CountRow As Integer
	Dim i As Integer
	Set rng = Selection
	
	CountRow = rng.EntireRow.Count
	For i = 1 To CountRow
		ActiveCell.EntireRow.Insert
		ActiveCell.Offset(2, 0).Select
	Next i
End Sub

```

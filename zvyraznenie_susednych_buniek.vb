```VB
' Vo výbere zvýrazni alternatívne riadky
' Highlight Alternate Rows in the Selection
Sub HighlightAlternateRows()
	Dim Myrange As Range
	Dim Myrow As Range
	Set Myrange = Selection
	
	For Each Myrow In Myrange.Rows
	   If Myrow.Row Mod 2 = 1 Then
		  Myrow.Interior.Color = vbCyan
	   End If
	Next Myrow
End Sub
```

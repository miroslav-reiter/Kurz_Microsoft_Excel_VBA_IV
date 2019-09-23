```VB
' Zvýraznite bunky chybne napísanými slovami
' Highlight Cells with Misspelled Words
Sub HighlightMisspelledCells()
	Dim cl As Range
	For Each cl In ActiveSheet.UsedRange
		If Not Application.CheckSpelling(word:=cl.Text) Then
			cl.Interior.Color = vbRed
		End If
	Next cl
End Sub

```

```VB
'Skryť všetky pracovné listy okrem aktívneho hárku
'Hide All Worksheets Except the Active Sheet

Sub HideAllExceptActiveSheet()
	Dim ws As Worksheet
	For Each ws In ThisWorkbook.Worksheets
		If ws.Name <> ActiveSheet.Name Then ws.Visible = xlSheetHidden
	Next ws
End Sub
```

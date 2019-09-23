Odkryt všetky pracovné hárky naraz
Unhide All Worksheets at One Go

```
' Tento kód skryje všetky hárky v zošite
' This code will unhide all sheets in the workbook
Sub UnhideAllWoksheets()
	Dim ws As Worksheet
	For Each ws In ActiveWorkbook.Worksheets
		ws.Visible = xlSheetVisible
	Next ws
End Sub
```

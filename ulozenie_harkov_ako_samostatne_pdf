```VB
' Uložte každý pracovný hárok ako samostatný súbor PDF
' Save Each Worksheet as a Separate PDF
Sub SaveWorkshetAsPDF()
	Dim ws As Worksheet
	For Each ws In Worksheets
		ws.ExportAsFixedFormat xlTypePDF, "C:\Test" & ws.Name & ".pdf"
	Next ws
	
' Dalsi Sposob
' ThisWorkbook.ExportAsFixedFormat xlTypePDF, "C:\Test" & ThisWorkbook.Name & ".pdf"
End Sub

```

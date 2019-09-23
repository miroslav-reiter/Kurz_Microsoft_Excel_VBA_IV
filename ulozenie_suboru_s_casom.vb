```
' Uloženie zošita s časovým údajom v jeho názve
' Save Workbook With TimeStamp in Its Name
Sub SaveWorkbookWithTimeStamp()
	Dim timestamp As String
	timestamp = Format(Date, "dd-mm-yyyy") & "_" & Format(Time, "hh-ss")
	ThisWorkbook.SaveAs "C:UsersUsernameDesktopWorkbookName" & timestamp
End Sub
```

```
' Zrušenie ochranu všetkých pracovných hárkov naraz
' Unprotect All Worksheets At One Go

Sub ProtectAllSheets()
	Dim ws As Worksheet
	Dim password As String
	password = "Testik123" 		'replace Testik123 with the password you want
	For Each ws In Worksheets
		ws.Unprotect password:=password
	Next ws
End Sub

```

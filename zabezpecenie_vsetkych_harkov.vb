```VB
' Zabezpečte naraz všetky pracovné hárky 
' Protect All Worksheets At One Go

'This code will protect all the sheets at one go
Sub ProtectAllSheets()
	Dim ws As Worksheet
	Dim password As String
	password = "Testik123" 	'replace Testik123 with the password you want
	
	For Each ws In Worksheets
	   ws.Protect password:=password
	Next ws
End Sub
```

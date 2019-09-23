```VB
' Ako získať iba číselnú časť z reťazca v Exceli
' How to Get Only the Numeric Part from a String in Excel
Function GetNumeric(CellRef As String)
	Dim StringLength As Integer
	StringLength = Len(CellRef)
	For i = 1 To StringLength
		If IsNumeric(Mid(CellRef, i, 1)) Then Result = Result & Mid(CellRef, i, 1)
	Next i
	GetNumeric = Result
End Function

```

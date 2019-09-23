```VB
' Zoradenie údajov podľa viacerých stĺpcov
' Sort Data by Multiple Columns
Sub SortMultipleColumns()
	With ActiveSheet.Sort
		 .SortFields.Add Key:=Range("A1"), Order:=xlAscending
		 .SortFields.Add Key:=Range("B1"), Order:=xlAscending
		 .SetRange Range("A1:C13")
		 .Header = xlYes
		 .Apply
	End With
End Sub

```

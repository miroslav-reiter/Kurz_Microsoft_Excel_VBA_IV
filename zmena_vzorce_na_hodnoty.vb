```VB
' Premeniť všetky vzorce na hodnoty
' Convert All Formulas into Values
Sub ConvertToValues()
	With ActiveSheet.UsedRange
		.Value = .Value
	End With
End Sub
```

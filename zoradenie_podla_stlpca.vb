```VB
' Ako triediť údaje podľa jedného stĺpca
' How to Sort Data by Single Column
Sub SortDataHeader()
	Range("DataRange").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
End Sub

```

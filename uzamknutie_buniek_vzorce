```VB
' Ochrana/Locknutie buniek so vzorcami a funkciami
' Protect/Lock Cells with Formulas
Sub LockCellsWithFormulas()
	With ActiveSheet
	   .Unprotect
	   .Cells.Locked = False
	   .Cells.SpecialCells(xlCellTypeFormulas).Locked = True
	   .Protect AllowDeletingRows:=True
	End With
End Sub

```

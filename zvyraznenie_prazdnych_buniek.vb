```VB
' Zvýraznite prázdne bunky pomocou VBA
' Highlight Blank Cells With VBA
Sub HighlightBlankCells()
	Dim Dataset as Range
	Set Dataset = Selection
	Dataset.SpecialCells(xlCellTypeBlanks).Interior.Color = vbRed
End Sub
``

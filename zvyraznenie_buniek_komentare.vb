```VB
' Zvýraznenie všetkých buniek s komentármi
' Highlight All Cells With Comments
Sub HighlightCellsWithComments()
	ActiveSheet.Cells.SpecialCells(xlCellTypeComments).Interior.Color = vbBlue
End Sub
``

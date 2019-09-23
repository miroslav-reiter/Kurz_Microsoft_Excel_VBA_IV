```VB
' Automaticky vložiť dátum a časovú pečiatku do susednej bunky
' Automatically Insert Date & Timestamp in the Adjacent Cell
Private Sub Worksheet_Change(ByVal Target As Range)
  On Error GoTo Handler
  If Target.Column = 1 And Target.Value <> "" Then
    Application.EnableEvents = False
    Target.Offset(0, 1) = Format(Now(), "dd-mm-yyyy hh:mm:ss")
    Application.EnableEvents = True
  End If
  Handler:
End Sub
```

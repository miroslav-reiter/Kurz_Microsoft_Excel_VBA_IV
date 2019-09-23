```VB
' Unprotect Worksheet
Sub Pwd()
On Error Resume Next
	For i = 65 To 66: For j = 65 To 66: For k = 65 To 66: For l = 65 To 66: For m = 65 To 66: For n = 65 To 66: For o = 65 To 66: For p = 65 To 66: For q = 65 To 66: For r = 65 To 66: For S = 65 To 66: For t = 32 To 126
	ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(n) & Chr(o) & Chr(p) & Chr(q) & Chr(r) & Chr(S) & Chr(t)
	If ActiveSheet.ProtectContents = False Then
		MsgBox “Worksheet unlocked, relock with: ” & Chr(i) & Chr(j) & Chr(k) & Chr(l) & ” ” & Chr(m) & Chr(n) & Chr(o) & Chr(p) & ” ” & Chr(q) & Chr(r) & Chr(S) & Chr(t)
		Exit Sub
	End If
	Next: Next: Next: Next: Next: Next: Next: Next: Next: Next: Next: Next
End Sub
``

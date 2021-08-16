'REM  *****  BASIC  *****
'LibreOffice Basic
Option Explicit

Private Sub main
	Dim msg As String
	msg = "test me"
    msg = includeHello(msg)
	Call printTarget(msg)
End Sub

Private Function includeHello(ByVal msg As String) As String
    includeHello = "Hello! " + msg
End Function

Private Sub printTarget(ByVal msg As String)
	msgbox(msg)
End Sub

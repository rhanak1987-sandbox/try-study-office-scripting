REM  *****  BASIC  *****
Option Explicit

' OEX = Using Option Explicit
' DIM = Declare Varialbe with Dim

' OEX | DIM | RES
'   0 |   0 |   1
'   0 |   1 |   1
'   1 |   0 |   0
'   1 |   1 |   1

'     | ... | DIM |
' ... |   1 |   1 |
' OEX |   0 |   1 |

Private Sub main
	Dim msg As String
	msg = "hello, test me"
	Call printTarget(msg)
End Sub

Private Sub printTarget(byval msg as string)
	msgbox(msg)
End Sub

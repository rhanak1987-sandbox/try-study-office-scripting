'REM  *****  BASIC  ***** (LibreOffice Basic)
'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8
Option Explicit

' OEX = Using Option Explicit
' DIM = Declare Variable with Dim

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

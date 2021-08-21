'REM  *****  BASIC  ***** tech: LibreOffice Basic

'Repo: https://github.com/rhanak1987-sandbox/try-study-office-scripting
'      as the developer and copyright holder (if applicable) of this project
'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8

'TODO: tidy with: https://github.com/todar/VBA-Style-Guide
'      as a user of the above style guide provided with "MIT License"
'look for: "TODO:", "FIXME:"
'tested: 21:18 2021-08-21 | ok
'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8

' OEX = 1 -> Using Option Explicit
' DIM = 1 -> Declare Variable with Dim
' RES = 0 -> runtime error: variable is not declared

' OEX | DIM | RES
'   0 |   0 |   1
'   0 |   1 |   1
'   1 |   0 |   0
'   1 |   1 |   1

'     | ... | DIM |
' ... |   1 |   1 |
' OEX |   0 |   1 |

Option Explicit

Private Sub main
    Dim msg As String
    msg = "hello, test me"
    Call printTarget(msg)
End Sub

Private Sub printTarget(byval msg as string)
    MsgBox(msg)
End Sub

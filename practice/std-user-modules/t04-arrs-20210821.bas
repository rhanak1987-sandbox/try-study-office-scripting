'REM  *****  BASIC  ***** tech: LibreOffice Basic

'Repo: https://github.com/rhanak1987-sandbox/try-study-office-scripting
'      as the developer and copyright holder (if applicable) of this project
'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8

'TODO: tidy with: https://github.com/todar/VBA-Style-Guide
'      as a user of the above style guide provided with "MIT License"
'look for: "TODO:", "FIXME:"
'tested: 19:14 2021-08-21 | ok

Option Explicit

Private Sub main
    DebugEmu.initDebug("")
    DebugEmu.pushToDebug("testing started: " & Now())

    Call testArrays

    DebugEmu.pushToDebug("testing finished: " & Now())
    Call printTarget(DebugEmu.getDebug)
End Sub

Private Sub testArrays
    'testing arrays and constants
    Dim debugMessage As String
    Const firstIndex As Long = 0

    Call addLineToDebug
    Dim fruits(4) As String
    fruits(0) = "apple"
    fruits(1) = "banana"
    fruits(2) = "citron"
    fruits(3) = "date"
    fruits(4) = "figs"
    debugMessage = fruits(firstIndex) & ", " & fruits(1)
    Call DebugEmu.pushToDebug(debugMessage)

    Call addLineToDebug
End Sub

Private Sub addLineToDebug
    Call DebugEmu.pushToDebug("- - - -")
End Sub

Private Sub printTarget(ByVal msg As String)
    Call MsgBox(msg)
End Sub

'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8
'use public Debug stack
'initDebug
'pushToDebug
'getDebug

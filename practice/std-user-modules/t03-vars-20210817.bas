'REM  *****  BASIC  ***** (LibreOffice Basic)
'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8
'tested: 20:34 2021-08-17 | ok
Option Explicit

Private Sub main
    DebugEmu.initDebug("")
    DebugEmu.pushToDebug("testing started: " & Now())

    Call testVariables

    DebugEmu.pushToDebug("testing finished: " & Now())
    Call printTarget(DebugEmu.getDebug)
End Sub

Private Sub testVariables
    Call addLineToDebug
    Dim int1 As Integer
    int1 = 32767 'int max.
    Dim int2%
    int2 = -32768 'int min.
    Dim int3 As Integer
    int3 = int1 + int2
    Call DebugEmu.pushToDebug(int3)

    Call addLineToDebug
    Dim lng1 As Long
    lng1 = 2147483647 'long max
    Dim lng2&
    lng2 = -2147483648 'long min
    Dim lng3 As Long
    lng3 = lng1 + lng2
    Call DebugEmu.pushToDebug(lng3)

    lng3 = lng1 / lng2
    Call DebugEmu.pushToDebug(lng3)
    'lng3 = lng1 + 1 'long_max + 1 -> overflow
    'Call DebugEmu.pushToDebug(lng3)

    Call addLineToDebug
End Sub

Private Sub addLineToDebug
    Call DebugEmu.pushToDebug("- - - -")
End Sub

Private Sub printTarget(ByVal msg As String)
    Call MsgBox(msg)
End Sub

'use public Debug stack
'initDebug
'pushToDebug
'getDebug

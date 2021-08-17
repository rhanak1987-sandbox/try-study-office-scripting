'REM  *****  BASIC  ***** tech: LibreOffice Basic

'Repo: https://github.com/rhanak1987-sandbox/try-study-office-scripting
'      as the developer and copyright holder (if applicable) of this project
'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8

'TODO: tidy with: https://github.com/todar/VBA-Style-Guide
'      as a user of the above style guide provided with "MIT License"
'look for: "TODO:", "FIXME:"
'tested: 23:18 2021-08-17 | ok

Option Explicit

Private Sub main
    DebugEmu.initDebug("")
    DebugEmu.pushToDebug("testing started: " & Now())

    Call testVariables

    DebugEmu.pushToDebug("testing finished: " & Now())
    Call printTarget(DebugEmu.getDebug)
End Sub

Private Sub testVariables
    Dim formattedNumber As String
    Dim debugMessage As String

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
    lng1 = 2147483647 'long max 2 147 483 647
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
    Dim sng1 As Single
    sng1 = 0.1
    Dim sng2!
    sng2 = 0.01
    Dim sng3 As Single
    sng3 = sng1 * sng2
    formattedNumber = Format(sng3, "0.####")
    debugMessage = "0.1(!) * 0.01(!) = " & formattedNumber & " = " & sng3
    Call DebugEmu.pushToDebug(debugMessage)
    sng3 = sng1 / sng2
    formattedNumber = Format(sng3, "0.####")
    debugMessage =  "0.1(!) / 0.01(!) = " & formattedNumber & " = " & sng3
    Call DebugEmu.pushToDebug(debugMessage)

    Call addLineToDebug
    Dim dbl1 As Double
    dbl1 = 0.1
    Dim dbl2!
    dbl2 = 0.01
    Dim dbl3 As Double
    dbl3 = dbl1 * dbl2
    formattedNumber = Format(dbl3, "0.####")
    debugMessage =  "0.1(#) * 0.01(#) = " & formattedNumber & " = " & dbl3
    Call DebugEmu.pushToDebug(debugMessage)
    dbl3 = fakeRound(dbl3, 10000)
    formattedNumber = Format(dbl3, "0.####")
    debugMessage =  "0.1(#) * 0.01(#) = " & formattedNumber & " = " & dbl3
    Call DebugEmu.pushToDebug(debugMessage)
    dbl3 = dbl1 / dbl2
    formattedNumber = Format(dbl3, "0.####")
    debugMessage =  "0.1(#) / 0.01(#) = " & formattedNumber & " = " & dbl3
    Call DebugEmu.pushToDebug(debugMessage)
    dbl3 = fakeRound(dbl3, 10000)
    formattedNumber = Format(dbl3, "0.####")
    debugMessage =  "0.1(#) / 0.01(#) = " & formattedNumber & " = " & dbl3
    Call DebugEmu.pushToDebug(debugMessage)

    Call addLineToDebug
    dbl3 = dbl1 / sng1
    formattedNumber = Format(dbl3, "0.####")
    debugMessage =  "0.1(#) / 0.1(!) = " & formattedNumber & " = " & dbl3
    Call DebugEmu.pushToDebug(debugMessage)
    dbl3 = fakeRound(dbl3, 10000)
    formattedNumber = Format(dbl3, "0.####")
    debugMessage =  "0.1(#) / 0.1(!) = " & formattedNumber & " = " & dbl3
    Call DebugEmu.pushToDebug(debugMessage)

    Call addLineToDebug
End Sub

Private Function fakeRound( _
	ByVal doubleVal As Double, _
	ByVal invPrecision As Long) _
	As Double
	'FIXME: This function is trouble, not a real rounding

    Dim longVal As Long
    longVal = doubleVal * invPrecision
    fakeRound = longVal / invPrecision
End Function

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

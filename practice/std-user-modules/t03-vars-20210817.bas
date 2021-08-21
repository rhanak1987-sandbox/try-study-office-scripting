'REM  *****  BASIC  ***** tech: LibreOffice Basic

'Repo: https://github.com/rhanak1987-sandbox/try-study-office-scripting
'      as the developer and copyright holder (if applicable) of this project
'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8

'TODO: tidy with: https://github.com/todar/VBA-Style-Guide
'      as a user of the above style guide provided with "MIT License"
'look for: "TODO:", "FIXME:"
'tested: 18:59 2021-08-21 | ok

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
    debugMessage = "32767(%) + -32768(%) = " & int3
    Call DebugEmu.pushToDebug(debugMessage)
    int3 = int1 - &h7FFF
    debugMessage = "32767(%) - &h7FFF = " & int3
    Call DebugEmu.pushToDebug(debugMessage)
    int3 = int1 - &o77777
    debugMessage = "32767(%) - &o77777 = " & int3
    Call DebugEmu.pushToDebug(debugMessage)

    Call addLineToDebug
    Dim lng1 As Long
    lng1 = 2147483647 'long max (2 147 483 647)
    Dim lng2&
    lng2 = -2147483648 'long min
    Dim lng3 As Long
    lng3 = lng1 + lng2
    debugMessage = "2147483647(&) + -2147483648(&) = " & lng3
    Call DebugEmu.pushToDebug(debugMessage)
    lng3 = lng1 / lng2
    debugMessage = "2147483647(&) / -2147483648(&) = " & lng3
    Call DebugEmu.pushToDebug(debugMessage)
    'lng3 = lng1 + 1 'long_max + 1 -> overflow
    'Call DebugEmu.pushToDebug(lng3)
    lng3 = lng1 - &h7fffffff
    debugMessage = "2147483647(&) - &h7fffffff = " & lng3
    Call DebugEmu.pushToDebug(debugMessage)

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
    Dim cur1 As Currency
    cur1 = 120
    Dim cur2@
    cur2 = 0.01
    Dim cur3 As Currency
    cur3 = cur1 * cur2
    formattedNumber = Format(cur3, "0.####")
    debugMessage =  "120(@) * 0.01(@) = " & formattedNumber & " = " & cur3
    Call DebugEmu.pushToDebug(debugMessage)

    Call addLineToDebug
    Dim str1 As String
    str1 = "apple"
    Dim str2$
    str2 = "banana"
    Dim str3 As String
    str3 = str1 & ", " & str2
    debugMessage = "str1($), str2($) = " & str3
    Call DebugEmu.pushToDebug(debugMessage)

    Call addLineToDebug
    Dim bln1 As Boolean
    bln1 = True
    Dim bln2 As Boolean
    bln2 = False
    Dim bln3 As Boolean
    bln3 = bln1 And bln2
    debugMessage = "True AND False = " & bln3
    Call DebugEmu.pushToDebug(debugMessage)
    bln3 = bln1 Or bln2
    debugMessage = "True OR False = " & bln3
    Call DebugEmu.pushToDebug(debugMessage)

    Call addLineToDebug
    Dim dte1 As Date
    Call DebugEmu.pushToDebug(dte1)
    dte1 = Now()
    Dim dte2 As Date
    dte2 = DateSerial(2021, 08, 21) + TimeSerial(18,30,00)
    Dim dte3 As Date
    dte3 = dte1 - dte2
    debugMessage = dte1 & ", " & dte2
    Call DebugEmu.pushToDebug(debugMessage)
    debugMessage = "diff = " & dte3
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

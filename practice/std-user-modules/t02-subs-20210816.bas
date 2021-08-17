'REM  *****  BASIC  ***** (LibreOffice Basic)
'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8
Option Explicit

Private Sub main
    DebugEmu.initDebug("")
    DebugEmu.pushToDebug("testing started")
    Call localTests
    DebugEmu.pushToDebug("testing finished")
    Call printTarget(DebugEmu.getDebug)
End Sub

Private Sub localTests
    Dim msg As String
    msg = "test me"

    msg = includeHello(msg) 'This calls a function
    Call addLineToDebug 'This calls a sub
    Call DebugEmu.pushToDebug(msg)
    Call addLineToDebug
    
    Call includeHelloVal(msg) 'ByVal param test
    Call DebugEmu.pushToDebug(msg)
    Call addLineToDebug
    
    Call includeHelloRef(msg) 'ByRef param test
    Call DebugEmu.pushToDebug(msg)
    Call addLineToDebug
End Sub

Private Function includeHello( _
    ByVal msg As String) _
    As String 'How to split long lines
    
    includeHello = "Hello! " + msg
End Function

Private Sub includeHelloVal(ByVal msg As String) As String
    msg = "Hello! " + msg
End Sub

Private Sub includeHelloRef(ByRef msg As String) As String
    msg = "Hello! " + msg
End Sub

Private Sub addLineToDebug
    Call DebugEmu.pushToDebug("- - - -")
End Sub

Private Sub printTarget(ByVal msg As String)
    msgbox(msg)
End Sub

'use public Debug stack
'initDebug
'pushToDebug
'getDebug

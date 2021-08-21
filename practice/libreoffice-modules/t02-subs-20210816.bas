'REM  *****  BASIC  ***** tech: LibreOffice Basic

'Repo: https://github.com/rhanak1987-sandbox/try-study-office-scripting
'      as the developer and copyright holder (if applicable) of this project
'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8

'TODO: tidy with: https://github.com/todar/VBA-Style-Guide
'      as a user of the above style guide provided with "MIT License"
'look for: "TODO:", "FIXME:"
'tested: 21:37 2021-08-21 | ok

Option Explicit

Private Sub main
    DebugEmu.initDebug("")
    DebugEmu.pushToDebug("testing started: " & Now())

    Call testSubroutines

    DebugEmu.pushToDebug("testing finished: " & Now())
    Call printTarget(DebugEmu.getDebug)
End Sub

Private Sub testSubroutines
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
    MsgBox(msg)
End Sub

'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8
'use public Debug stack
'initDebug
'pushToDebug
'getDebug

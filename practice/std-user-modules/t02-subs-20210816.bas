'REM  *****  BASIC  *****
'LibreOffice Basic
Option Explicit

'use public Debug stack
'initDebug
'pushToDebug
'getDebug

Private Sub main
	DebugEmu.initDebug("")
	DebugEmu.pushToDebug("testing started")
	Call localTests
	DebugEmu.pushToDebug("testing finished")
	Call printTarget(getDebug)
End Sub

Private Sub localTests
	Dim msg As String
	msg = "test me"
    msg = includeHello(msg) 'This calls a function
    Call addLineToDebug 'This calls a sub
	Call DebugEmu.pushToDebug(msg)
	Call addLineToDebug
End Sub

Private Function includeHello(ByVal msg As String) As String
    includeHello = "Hello! " + msg
End Function

Private Sub addLineToDebug
	Call DebugEmu.pushToDebug("- - - -")
End Sub

Private Sub printTarget(ByVal msg As String)
	msgbox(msg)
End Sub

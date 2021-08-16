'REM  *****  BASIC  ***** (LibreOffice Basic)
'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8
Option Explicit

'use public Debug stack
'initDebug
'pushToDebug
'getDebug

Private debugString As String 'can be watched during debug

Private Sub main
    'testing debug emu
    debugString = "init"
    msgbox(debugString)
    
    initDebug("")
    pushToDebug("Hello")
    pushToDebug("World")
    pushToDebug("2021")
    msgbox(getDebug)
End Sub

Public Sub initDebug(ByVal debugText As String)
    debugString = ""
End Sub

Public Sub pushToDebug(ByVal debugText As String)
    Dim separator As String 'vbNewLine, vbCrLf, vbCr, vbLf
    separator = Chr(10)
    
    debugString = debugString & debugText & separator
End Sub

Public Function getDebug As String
    getDebug = debugString
End Function

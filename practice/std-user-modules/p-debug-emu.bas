'REM  *****  BASIC  ***** tech: LibreOffice Basic

'Repo: https://github.com/rhanak1987-sandbox/try-study-office-scripting
'      as the developer and copyright holder (if applicable) of this project
'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8

'TODO: tidy with: https://github.com/todar/VBA-Style-Guide
'      as a user of the above style guide provided with "MIT License"
'look for: "TODO:", "FIXME:"
'tested: 23:19 2021-08-17 | ok

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

    initDebug("start over")
    msgbox(getDebug)
End Sub

Public Sub initDebug(ByVal debugText As String)
    debugString = debugText
End Sub

Public Sub pushToDebug(ByVal debugText As String)
    Dim separator As String 'vbNewLine, vbCrLf, vbCr, vbLf
    separator = Chr(10)
    
    debugString = debugString & debugText & separator
End Sub

Public Function getDebug As String
    getDebug = debugString
End Function

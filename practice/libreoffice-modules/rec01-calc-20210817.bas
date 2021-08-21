'REM  *****  BASIC  ***** (LibreOffice Basic)
'...;....1....;....2....;....3....;....4....;....5....;....6....;....7....;....8
' Takeaway: "why not to use macro recorder"

Sub Main

End Sub


sub Macro1
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
dim args1(0) as new com.sun.star.beans.PropertyValue
args1(0).Name = "StringName"
args1(0).Value = "1"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args1())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextCell", "", 0, Array())

rem ----------------------------------------------------------------------
dim args3(0) as new com.sun.star.beans.PropertyValue
args3(0).Name = "StringName"
args3(0).Value = "2"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args3())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextCell", "", 0, Array())

rem ----------------------------------------------------------------------
dim args5(0) as new com.sun.star.beans.PropertyValue
args5(0).Name = "ToPoint"
args5(0).Value = "$A$1:$A$2"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args5())

rem ----------------------------------------------------------------------
dim args6(0) as new com.sun.star.beans.PropertyValue
args6(0).Name = "EndCell"
args6(0).Value = "$A$10"

dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, args6())

rem ----------------------------------------------------------------------
dim args7(0) as new com.sun.star.beans.PropertyValue
args7(0).Name = "ToPoint"
args7(0).Value = "$A$1:$A$10"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args7())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:InsertRowsBefore", "", 0, Array())

rem ----------------------------------------------------------------------
dim args9(0) as new com.sun.star.beans.PropertyValue
args9(0).Name = "ToPoint"
args9(0).Value = "$A$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args9())

rem ----------------------------------------------------------------------
dim args10(0) as new com.sun.star.beans.PropertyValue
args10(0).Name = "StringName"
args10(0).Value = "id"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args10())

rem ----------------------------------------------------------------------
dim args11(0) as new com.sun.star.beans.PropertyValue
args11(0).Name = "StringName"
args11(0).Value = "val"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args11())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextCell", "", 0, Array())

rem ----------------------------------------------------------------------
dim args13(1) as new com.sun.star.beans.PropertyValue
args13(0).Name = "By"
args13(0).Value = 1
args13(1).Name = "Sel"
args13(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args13())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextCell", "", 0, Array())

rem ----------------------------------------------------------------------
dim args15(0) as new com.sun.star.beans.PropertyValue
args15(0).Name = "ToPoint"
args15(0).Value = "$B$2"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args15())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, Array())

rem ----------------------------------------------------------------------
dim args17(0) as new com.sun.star.beans.PropertyValue
args17(0).Name = "ToPoint"
args17(0).Value = "$B$2"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args17())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, Array())

rem ----------------------------------------------------------------------
dim args19(0) as new com.sun.star.beans.PropertyValue
args19(0).Name = "ToPoint"
args19(0).Value = "$B$2:$B$11"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args19())

rem ----------------------------------------------------------------------
dim args20(0) as new com.sun.star.beans.PropertyValue
args20(0).Name = "ToPoint"
args20(0).Value = "$A$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args20())

rem ----------------------------------------------------------------------
dim args21(1) as new com.sun.star.beans.PropertyValue
args21(0).Name = "By"
args21(0).Value = 1
args21(1).Name = "Sel"
args21(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoDownToEndOfData", "", 0, args21())

rem ----------------------------------------------------------------------
dim args22(1) as new com.sun.star.beans.PropertyValue
args22(0).Name = "By"
args22(0).Value = 1
args22(1).Name = "Sel"
args22(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRightToEndOfData", "", 0, args22())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:SelectAll", "", 0, Array())

rem ----------------------------------------------------------------------
dim args24(1) as new com.sun.star.beans.PropertyValue
args24(0).Name = "By"
args24(0).Value = 1
args24(1).Name = "Sel"
args24(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoUp", "", 0, args24())

rem ----------------------------------------------------------------------
dim args25(1) as new com.sun.star.beans.PropertyValue
args25(0).Name = "By"
args25(0).Value = 1
args25(1).Name = "Sel"
args25(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoUpToStartOfData", "", 0, args25())

rem ----------------------------------------------------------------------
dim args26(1) as new com.sun.star.beans.PropertyValue
args26(0).Name = "By"
args26(0).Value = 1
args26(1).Name = "Sel"
args26(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeftToStartOfData", "", 0, args26())

rem ----------------------------------------------------------------------
dim args27(0) as new com.sun.star.beans.PropertyValue
args27(0).Name = "By"
args27(0).Value = 1

dispatcher.executeDispatch(document, ".uno:GoRightToEndOfDataSel", "", 0, args27())

rem ----------------------------------------------------------------------
dim args28(0) as new com.sun.star.beans.PropertyValue
args28(0).Name = "By"
args28(0).Value = 1

dispatcher.executeDispatch(document, ".uno:GoDownToEndOfDataSel", "", 0, args28())

rem ----------------------------------------------------------------------
dim args29(1) as new com.sun.star.beans.PropertyValue
args29(0).Name = "By"
args29(0).Value = 1
args29(1).Name = "Sel"
args29(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoDownToEndOfData", "", 0, args29())

rem ----------------------------------------------------------------------
dim args30(1) as new com.sun.star.beans.PropertyValue
args30(0).Name = "By"
args30(0).Value = 1
args30(1).Name = "Sel"
args30(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoUpToStartOfData", "", 0, args30())

rem ----------------------------------------------------------------------
dim args31(1) as new com.sun.star.beans.PropertyValue
args31(0).Name = "By"
args31(0).Value = 1
args31(1).Name = "Sel"
args31(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args31())

rem ----------------------------------------------------------------------
dim args32(0) as new com.sun.star.beans.PropertyValue
args32(0).Name = "StringName"
args32(0).Value = "11"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args32())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextCell", "", 0, Array())

rem ----------------------------------------------------------------------
dim args34(1) as new com.sun.star.beans.PropertyValue
args34(0).Name = "By"
args34(0).Value = 1
args34(1).Name = "Sel"
args34(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoUp", "", 0, args34())

rem ----------------------------------------------------------------------
dim args35(1) as new com.sun.star.beans.PropertyValue
args35(0).Name = "By"
args35(0).Value = 1
args35(1).Name = "Sel"
args35(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args35())

rem ----------------------------------------------------------------------
dim args36(1) as new com.sun.star.beans.PropertyValue
args36(0).Name = "By"
args36(0).Value = 1
args36(1).Name = "Sel"
args36(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoUp", "", 0, args36())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())

rem ----------------------------------------------------------------------
dim args38(1) as new com.sun.star.beans.PropertyValue
args38(0).Name = "By"
args38(0).Value = 1
args38(1).Name = "Sel"
args38(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args38())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

rem ----------------------------------------------------------------------
dim args40(1) as new com.sun.star.beans.PropertyValue
args40(0).Name = "By"
args40(0).Value = 1
args40(1).Name = "Sel"
args40(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args40())

rem ----------------------------------------------------------------------
dim args41(1) as new com.sun.star.beans.PropertyValue
args41(0).Name = "By"
args41(0).Value = 1
args41(1).Name = "Sel"
args41(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoUpToStartOfData", "", 0, args41())

rem ----------------------------------------------------------------------
dim args42(1) as new com.sun.star.beans.PropertyValue
args42(0).Name = "By"
args42(0).Value = 1
args42(1).Name = "Sel"
args42(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args42())

rem ----------------------------------------------------------------------
dim args43(1) as new com.sun.star.beans.PropertyValue
args43(0).Name = "By"
args43(0).Value = 1
args43(1).Name = "Sel"
args43(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args43())

rem ----------------------------------------------------------------------
dim args44(0) as new com.sun.star.beans.PropertyValue
args44(0).Name = "StringName"
args44(0).Value = "=VÉL()*10"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args44())

rem ----------------------------------------------------------------------
dim args45(1) as new com.sun.star.beans.PropertyValue
args45(0).Name = "By"
args45(0).Value = 1
args45(1).Name = "Sel"
args45(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoUp", "", 0, args45())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, Array())

rem ----------------------------------------------------------------------
dim args47(0) as new com.sun.star.beans.PropertyValue
args47(0).Name = "ToPoint"
args47(0).Value = "$B$2"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args47())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, Array())

rem ----------------------------------------------------------------------
dim args49(0) as new com.sun.star.beans.PropertyValue
args49(0).Name = "ToPoint"
args49(0).Value = "$B$2:$B$12"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args49())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:NumberFormatDecimal", "", 0, Array())

rem ----------------------------------------------------------------------
dim args51(0) as new com.sun.star.beans.PropertyValue
args51(0).Name = "ToPoint"
args51(0).Value = "$A$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args51())

rem ----------------------------------------------------------------------
dim args52(1) as new com.sun.star.beans.PropertyValue
args52(0).Name = "By"
args52(0).Value = 1
args52(1).Name = "Sel"
args52(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoDownToEndOfData", "", 0, args52())

rem ----------------------------------------------------------------------
dim args53(1) as new com.sun.star.beans.PropertyValue
args53(0).Name = "By"
args53(0).Value = 1
args53(1).Name = "Sel"
args53(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args53())

rem ----------------------------------------------------------------------
dim args54(0) as new com.sun.star.beans.PropertyValue
args54(0).Name = "StringName"
args54(0).Value = "12"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args54())

rem ----------------------------------------------------------------------
dim args55(1) as new com.sun.star.beans.PropertyValue
args55(0).Name = "By"
args55(0).Value = 1
args55(1).Name = "Sel"
args55(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args55())

rem ----------------------------------------------------------------------
dim args56(1) as new com.sun.star.beans.PropertyValue
args56(0).Name = "By"
args56(0).Value = 1
args56(1).Name = "Sel"
args56(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoUp", "", 0, args56())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())

rem ----------------------------------------------------------------------
dim args58(1) as new com.sun.star.beans.PropertyValue
args58(0).Name = "By"
args58(0).Value = 1
args58(1).Name = "Sel"
args58(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args58())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())


end sub
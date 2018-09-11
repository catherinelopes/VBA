Attribute VB_Name = "Module1"
Option Explicit

' This is a macro for multiply rows with user defined number '

Sub MultiplyRows()
Dim RwsCnt As Long, LR As Long, InsRw As Long


RwsCnt = Application.InputBox("How many copies of each row should be inserted?", "Insert Count", 2, Type:=1)
If RwsCnt = 0 Then Exit Sub
LR = Range("A" & Rows.Count).End(xlUp).Row

Application.ScreenUpdating = False
' from row 2 to the last row, exclude the header'
For InsRw = LR To 2 Step -1
    Rows(InsRw).Copy
    Rows(InsRw + 1).Resize(RwsCnt).Insert xlShiftDown
Next InsRw
Application.ScreenUpdating = True

End Sub


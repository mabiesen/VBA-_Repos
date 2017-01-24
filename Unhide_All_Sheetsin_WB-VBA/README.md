# UnhideAllSheetsinWB
To do a comprehensive "find" in a workbook, all sheets must be unhidden.  Excel offers no easy way to unhide all sheets at once, this macro performs that task.

```vba

Sub UnhideAllSheets()
    Dim ws As Worksheet
 Ans = MsgBox("You have selected workbook " & ActiveCell.Parent.Parent.Name & _
 ". Is this correct?", vbYesNo + vbQuestion, "Request Summary")
If Ans = vbNo Then Exit Sub
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
 
End Sub

```

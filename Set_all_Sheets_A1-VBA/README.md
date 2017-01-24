# SetAllSheetsA1-VBA
This macro is used in multi-spreadsheet reports to insure that users opening the file for the first time are directed to cell A1.  Reduces user confusion.

```vba
Sub SetALLSheetsA1()
Dim ws As Worksheet
Ans = MsgBox("Hi " & Environ("Username") & ". You have selected workbook " & ActiveCell.Parent.Parent.Name & _
". Is this correct?", vbYesNo + vbQuestion, "Request Summary")
If Ans = vbNo Then Exit Sub
    For Each ws In ActiveWorkbook.Sheets
        ws.Activate
        ws.[a1].Select
    Next ws
    ActiveWorkbook.Worksheets(1).Activate
End Sub
```

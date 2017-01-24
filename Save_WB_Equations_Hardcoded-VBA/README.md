# SaveWBEquationsHardcoded
Converting equations to numeric figures decreases file size and preserves data historically.  That is the intent of this macro.

```vba

Sub savehardcoded()
Dim file_name As Variant
Dim sh As Worksheet
'save this workbook
ThisWorkbook.Save
'save as to new file
file_name = Application.GetSaveAsFilename(FileFilter:="Microsoft Excel file (*.xls), *.xls")
    If file_name <> False Then
      ActiveWorkbook.SaveAs Filename:=file_name
      MsgBox "File Saved!"
    End If

ThisWorkbook.Activate
'unhide all sheets
    For Each sh In ActiveWorkbook.Worksheets
        sh.Visible = xlSheetVisible
    Next sh
    
    For Each sh In ThisWorkbook.Worksheets
        If sh.Visible = True Then
            sh.Activate
            sh.Cells.Copy
            sh.Range("A1").PasteSpecial Paste:=xlValues
            sh.Range("A1").Select
        End If
    Next sh
    Application.CutCopyMode = False

End Sub

```

# FillBlankCells-Excel
Code to quickly fill blank cells in preparation for other tasks.  Created to avoid issues with null values in Access.

```

Sub FillEmptyBlankCellWithValue()
Dim cell As Range
Dim InputValue As String
On Error Resume Next
InputValue = InputBox("Enter value that will fill empty cells in selection", _
"Fill Empty Cells")
For Each cell In Selection
If IsEmpty(cell) Then
cell.Value = InputValue
End If
Next
End Sub

```

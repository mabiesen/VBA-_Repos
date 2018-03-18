## Fetch comments from a given cell
```
Function get_comment_from_cell(this_sheet As Worksheet, cell_row As Long, cell_column As Long) As String

    If Not this_sheet.Cells(cell_row, cell_column).Comment Is Nothing Then
        s = this_sheet.Cells(cell_row, cell_column).Comment.Text
    Else
        s = ""
    End If
    
    get_comment_from_cell = s

End Function
```

## Add comment to a given cell
```
Sub add_comment_to_cell(this_sheet As Worksheet, this_comment As String, cell_row As Long, cell_column As Long)

    'Convert to range string
    this_sheet.Cells(cell_row, cell_column).AddComment this_comment

End Sub
```


## Remove comments from a given range of cells
```
Sub remove_comment_from_cell(this_sheet As Worksheet, cell_row As Long, cell_column As Long)
    
    Dim thisthing As String
    
    thisthing = get_comment_from_cell(this_sheet, cell_row, cell_column)
    If Not thisthing = "" Then
        this_sheet.Cells(cell_row, cell_column).ClearComments
    End If

End Sub
```

## Copy comment to new cell.  Uses two of the above functions
```
Sub copy_comment_to_new_cell(orig_sheet As Worksheet, new_sheet As Worksheet, orig_cell_row As Long, orig_cell_column As Long, new_cell_row As Long, new_cell_column As Long)
    
    Dim thisthing As String
    thisthing = get_comment_from_cell(orig_sheet, orig_cell_row, orig_cell_column)
    If Not thisthing = "" Then
        add_comment_to_cell new_sheet, thisthing, new_cell_row, new_cell_column
    End If

End Sub

```

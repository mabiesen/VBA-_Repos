## Set subtotals as last row of data for given range
```
Sub update_subtotals(this_sheet As Worksheet, this_row As Long, num_rows As Long, start_col As Long, end_col As Long)
    'Tested
    Dim i As Long
    Dim col_letter As String
    
    For i = start_col To end_col
        col_letter = util_cols.col_letter(i)
        this_sheet.Cells(this_row, i).Value = "=sum(" & col_letter & (this_row - num_rows) & ":" & col_letter & (this_row - 1) & ")"
    Next i

End Sub
```

## Set row values in column range
```
Sub cell_values_to_zero(this_sheet As Worksheet, this_row As Long, start_col As Long, end_col As Long, cell_content as Variant)
    'Tested
    Dim i As Long
    
    For i = start_col To end_col
        this_sheet.Cells(this_row, i).Value = cell_content
    Next i

End Sub
```

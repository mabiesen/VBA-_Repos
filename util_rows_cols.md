## Check a range in a column for blanks, return true/false
```
Function check_range_in_column_for_blanks(this_sheet As Worksheet, col_num As Long, start_row As Long, end_row As Long) As Boolean
    'Tested
    Dim ctr As Long
    Dim return_value As Boolean
    
    For ctr = start_row To end_row
        If this_sheet.Cells(ctr, col_num).Value = "" Then
            return_value = True
        End If
    Next ctr
    
    If return_value = True Then
        check_range_in_column_for_blanks = True
    Else
        check_range_in_column_for_blanks = False
    End If

End Function
```

## Check range in column for dupes, return true/false
```
Function check_dupes_in_column(this_sheet As Worksheet, col_num As Long, start_row As Long, end_row As Long) As Boolean

    Dim return_value As Boolean
    Dim matchFoundIndex As Long
    Dim iCntr As Long
    
    For iCntr = start_row To end_row
        If Cells(iCntr, 1) <> "" Then
            matchFoundIndex = WorksheetFunction.Match(Cells(iCntr, 1), Range("A1:A" & end_row), 0)
            If iCntr <> matchFoundIndex Then
                return_value = True
            End If
        End If
    Next
    
    If return_value = True Then
        check_dupes_in_column = True
    Else
        check_dupes_in_column = False
    End If

End Function
```
## Return number of first blank row
```
Function get_first_sheet_blank_row(this_sheet As Worksheet) As Long
    'Tested
    Dim num_lrow As Long
    
    num_lrow = this_sheet.Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    get_first_sheet_blank_row = num_lrow

End Function
```

## Return number of last data row
```
Function get_last_sheet_data_row(this_sheet As Worksheet) As Long
    'Tested
    Dim num_lrow As Long
    
    num_lrow = this_sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    get_last_sheet_data_row = num_lrow

End Function
```

## Return number of active rows (rows containing data for a given column)
```
Function get_num_active_rows(this_sheet As Worksheet, num_startrow As Long, num_col As Long, end_indicator As String) As Long
    'Tested
    Dim i As Long
    
    For i = num_startrow To this_sheet.Rows.Count
        
        If this_sheet.Cells(i, num_col).Value = end_indicator Then
            get_num_active_rows = i - num_startrow
            Exit For
        End If
        
    Next i


End Function
```

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

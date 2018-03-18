## Set or remove protections for sheet.  True sets pw
```
Sub lock_or_unlock_sheet(this_sheet As Worksheet, this_bool As Boolean)

    If this_bool = True Then
        this_sheet.Protect Password:="haha2626"
    Else
        this_sheet.Unprotect Password:="haha2626"
    End If
    
End Sub
```

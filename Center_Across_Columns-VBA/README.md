# CenterAcrossColumns-MergeCenterAlt
Merge and center can be very disruptive to Excel equations.  A better alternative is horizonal alignment.

```vba
Sub CenterAcrossColumns()
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .MergeCells = False
    End With
End Sub
```

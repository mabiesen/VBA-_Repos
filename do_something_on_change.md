
This sub module will perform some action when data some portion of the worksheet has changed

To use, place this submodule inside of the worksheet object.

```
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$B$2" Then
        '
        'Enter Code or Call any Function if any process has to be performed
        'When someone Edits the cell B2
        '
        '
    End If
End Sub
```

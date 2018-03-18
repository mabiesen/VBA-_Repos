```
Function col_letter(lngCol As Long) As String
    'Tested
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    col_letter = vArr(0)
End Function
```

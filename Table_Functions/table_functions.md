# Table Functions

## 1. Super Basic Functions

```

' provide worksheet object. get table name as string.  good if only one table in sheet with default name
Public Function get_table_name_one_table_in_sheet(ByRef mysheet As Worksheet) As String
    Dim tablename As String
    tablename = mysheet.ListObjects(1).Name
    Debug.Print (tablename)
    get_table_name_one_table_in_sheet = tablename
End Function



' Find the number of columns in our table
Public Function get_table_width(ByRef table As ListObject) As Integer
    Dim numcol As Integer
    
    numcol = table.HeaderRowRange.Count
    
    get_table_width = numcol

End Function

' Find the number of rows in our table
Public Function get_table_height(ByRef table As ListObject) As Long
    Dim numrows As Long
    
    numrows = table.DataBodyRange.Rows.Count
    
    get_table_height = numrows

End Function

```

## 2. Manage Table Data With Arrays

```
' MODULE PURPOSE: FUNCTIONS FOR TABLE ARRAY CREATION AND MANIPULATION

' provide a table object, return an array of rows
' Note: this function is dependent upon other functions in the basic_table_funcs module
Public Function get_table_array_of_rows(ByRef table As ListObject) As Variant
    Dim numcols As Long
    Dim numrows As Long
    Dim mytabledata() As Variant
    
    numcols = get_table_width(table)
    numrows = get_table_height(table)
    mytabledata = get_table_body_data(table)
    
    get_table_array_of_rows = twod_array_from_one(mytabledata, numcols, numrows)
    
End Function

' provide worksheet object. find all table names in sheet.  Return as array of strings
Public Function get_table_name_all_tables_in_sheet(ByRef mysheet As Worksheet) As Variant
    Dim tbl As ListObject
    Dim numtables As Integer
    Dim myarray() As String

    numtables = 0
    
    For Each tbl In mysheet.ListObjects
        ReDim Preserve myarray(numtables)
        myarray(numtables) = tbl.Name
        numbtables = numtables + 1
    Next tbl
    
    get_table_name_all_tables_in_sheet = myarray
    
End Function

' Provide table object.  Obtain headers, return as array of strings
Public Function get_table_headers(ByRef tableobj As ListObject) As Variant
    Dim objListRng As Range
    Dim cell As Range
    Dim myarray() As String
    Dim numcol As Integer
    
    objListRng = tableobj.HeaderRowRange
    numcol = 0
    
    For Each cell In objListRng
        ReDim Preserve myarray(numcol)
        myarray(numcol) = cell.Value
        numcol = numcol + 1
    Next cell
    
    get_table_headers = myarray

End Function

' Provide table object. Create databody array. Return as variant array of all data
Public Function get_table_body_data(ByRef table As ListObject) As Variant
    Dim myrange As Range
    Dim cell As Range
    Dim myarray() As Variant
    Dim numcells As Long
    
    numcells = 0
    Set myrange = table.DataBodyRange
    
    For Each cell In myrange
        ReDim Preserve myarray(numcells)
        myarray(numcells) = cell.Value
        numcells = numcells + 1
    Next cell
    
    get_table_body_data = myarray

End Function

```

## 3. Table Sum Function

```
'Function aims to sum provided columns in table as a singular sum
'Column criteria is provided to the function as an array of column indexes
'myarray is a two dimensional array consisiting of table data.  It is an array of rows so to speak

Public Function sum_twod_array(ByRef myarray As Variant, ByRef criteriaarray As Variant) As Long
Dim firstdimlength As Long
Dim seconddimlength As Long
Dim criterialength As Long
Dim rowctr As Long
Dim colctr As Long
Dim mysum As Long

firstdimlength = UBound(myarray, 1)
seconddimlength = UBound(myarray, 2)
criterialength = UBound(criteriaarray, 1)

rowctr = 0
colctr = 0
mysum = 0

Do While rowctr < firstdimlength
    Do While colctr < seconddimlength
        If IsInArray(colctr, criteriaarray) Then
            mysum = mysum + myarray(rowctr, colctr)
        End If
        colctr = colctr + 1
    Loop
    colctr = 0
    rowctr = rowctr + 1
Loop

sum_twod_array = mysum

End Function
```

## 4. Array Functions Required for the Above Table Functions

```
'Determine if item is inside of array
Function IsInArray(valueToBeFound As Variant, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, valueToBeFound)) > -1)
End Function

' non-specific multidimensional array function
Public Function twod_array_from_one(ByRef initialarray() As Variant, numcols As Long, numrows As Long) As Variant

    Dim masterarray() As Variant
    Dim colctr As Long
    Dim rowctr As Long
    Dim masterctr As Long

    colctr = 0
    rowctr = 0
    masterctr = 0
    
    ReDim masterarray(numrows, numcols) As Variant
    
    Do While rowctr < numrows
    
        Do While colctr < numcols
            masterarray(rowctr, colctr) = initialarray(masterctr)
            colctr = colctr + 1
            masterctr = masterctr + 1
        Loop
        colctr = 0
        rowctr = rowctr + 1
    Loop
    
    twod_array_from_one = masterarray

End Function



```

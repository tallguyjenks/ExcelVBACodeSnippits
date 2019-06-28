Dim ws As Worksheet
    Set ws = ActiveSheet
Dim tbl As ListObject
	'you input the name of the table
    Set tbl = ws.ListObjects("categories")
Dim sortcolumn As Range
    'name of the table and then column header in the brackets
    Set sortcolumn = Range("categories[Count]")
        With tbl.Sort
           .SortFields.Clear
           .SortFields.Add _
                Key:=sortcolumn, _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending
           .Header = xlYes
           .Apply
        End With

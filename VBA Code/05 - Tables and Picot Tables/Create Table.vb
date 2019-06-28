'CREATE TABLE WITH INPUTTED BOUNDRIES AND NAME IT DATATABLE
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(LastRow, LastColumn)), , xlYes).Name = "DataTable"
   
    
 Dim ws As Worksheet, tbl As ListObject, sortcolumn As Range
 Dim DateNum As Long, Ordertype As String
            Set ws = ActiveSheet
            Set tbl = ws.ListObjects("DataTable")
            Set sortcolumn = Range("DataTable[Order type]")

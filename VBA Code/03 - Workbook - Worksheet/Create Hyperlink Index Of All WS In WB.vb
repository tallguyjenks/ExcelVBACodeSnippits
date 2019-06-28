Sub ListSheets()
    Dim ws As Worksheet
    Dim x As Integer
    
    x = 1
    
    ActiveSheet.Range("A:A").Clear
    
    For Each ws In Worksheets
        
        ActiveSheet.Cells(x, 1).Select
        
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=vbNullString, SubAddress:=ws.Name & "!A1", TextToDisplay:=ws.Name
        
        x = x + 1
        
    Next ws
    
End Sub

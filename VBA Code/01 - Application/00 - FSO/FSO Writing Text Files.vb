Sub creatinganewtextfile()
    Dim FSO As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
        Set FSO = New Scripting.FileSystemObject
        Set ts = FSO.CreateTextFile(Environ("userprofile") & "\desktop\" & "activity log.txt")

    ts.Write "created on " & Now & vbNewLine
    ts.WriteLine "Created by " & Environ("username")
    ts.WriteBlankLines 2
    ts.Write "Data starts here" & vbNewLine

    ts.Close
    Set FSO = Nothing
    Call adddatatotextfile
End Sub

Sub adddatatotextfile()
    Dim FSO As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim r As Range
    Dim colcount As Integer
    Dim i As Integer
        Set FSO = New Scripting.FileSystemObject
        Set ts = FSO.OpenTextFile(Environ("userprofile") & "\desktop\" & "activity log.txt", ForAppending)
        
        Sheets("vba playground").Activate
        
        colcount = Range("a1", Range("a1").End(xlToRight)).Cells.Count
        
        For Each r In Range("A2", Range("a1").End(xlDown))
        
            For i = 1 To colcount
                
                ts.Write r.Offset(0, i - 1).Value
                If i < colcount Then ts.Write vbTab
                
            Next i
            
            ts.WriteLine
            
        Next r
        
        ts.Close
        Set FSO = Nothing
        

End Sub



==========================================
write values from texst file to excel
==========================================

Sub readfromtextfile()
    Dim FSO As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    Dim textline As String
    Dim tabposition As Integer
        Set FSO = New Scripting.FileSystemObject
        Set ts = FSO.OpenTextFile(Environ("userprofile") & "\desktop\" & "activity log.txt")
Application.ScreenUpdating = False
        Workbooks.Add
        
        Do Until ts.AtEndOfStream
            textline = ts.ReadLine
            tabposition = InStr(textline, vbTab)
            Do Until tabposition = 0
                activecell.Value = Left(textline, tabposition - 1)
                activecell.Offset(0, 1).Select
                textline = Right(textline, Len(textline) - tabposition)
                tabposition = InStr(textline, vbTab)
            Loop
            activecell.Value = textline
            activecell.Offset(1, 0).End(xlToLeft).Select
        Loop

        ts.Close
        Set FSO = Nothing
Application.ScreenUpdating = True
End Sub



==========================================

Dim strFileName As String, myFile As Integer, strText As String

Close #myFile
==========================================
strFileName = "C:\test.txt"
myFile = FreeFile 'Get first free file number
==========================================
'Open Text file for Input to extract information
Open strFileName For Input As #myFile

strFileName = "C:\test.txt"
myFile = FreeFile 'Get first free file number
==========================================
'Open Text file for Output to Write to the File
Open strFileName For Output As #myFile

strText = Input$(LOF(myFile), myFile)

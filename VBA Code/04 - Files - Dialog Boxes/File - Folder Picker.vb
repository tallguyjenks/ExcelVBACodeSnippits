Sub FilePicker()
Dim fd As FileDialog
Dim filewaschosen As Boolean
    Set fd = Application.FileDialog(msoFileDialogOpen)

'set your own filters for file types you want shown
    fd.Filters.Clear
        fd.Filters.Add "old excel files", "*.xls"
        fd.Filters.Add "New excel files", "*.xlsx"
        fd.Filters.Add "Macro excel files", "*.xlsm"
        fd.Filters.Add "ALL excel files", "*.xl*"
    fd.FilterIndex = 4


'doesnt allow multip file selections in the dialog box
    fd.AllowMultiSelect = False

'initial filename where the dialog box opens up upon
    fd.InitialFileName = Environ("userprofile") & "\desktop"

'changes the title of the dialog box header
    fd.Title = "open sesame!"
'updates to this setting upon selecting a file to open
    fd.ButtonName = "GO!"

'file: picker will return the path for the file
'folder: picker will return the path for the folder
'open: open a file
'saveas: a file

filewaschosen = fd.Show
'open = 1 close = 0
If Not filewaschosen Then
    MsgBox "you didnt select a file"
        Exit Sub
End If

fd.Execute 'executes the default action
End Sub

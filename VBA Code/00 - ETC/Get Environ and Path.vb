Dim WshShell As Object
Dim SpecialPath As String

Set WshShell = CreateObject("WScript.Shell")
SpecialPath = WshShell.SpecialFolders("Desktop")

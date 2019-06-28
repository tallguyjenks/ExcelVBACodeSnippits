'CREATING THE INSTANCE OF THE FILE SYSTEM OBJECT


Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")


'3 GET METHODS OF THE FSO:
'GetDrive
'GetFolder
	'GetParentFolderName 'Determine the name of a folder's parent folder.
	'GetSpecialFolder 'Determine the path of system folders.
'GetFile


'OTHER FSO METHODS
'CreateFolder

'File.Move or FileSystemObject.MoveFile
'File.Copy or FileSystemObject.CopyFile
'File.Delete or FileSystemObject.DeleteFile

'FolderExists
'DateLastModified



Dim fso, f1
Set fso = CreateObject("Scripting.FileSystemObject")
Set f1 = fso.GetFile("c:\test.txt")


'syntax to create a folder and get a response about it


Sub CreateFolder
   Dim fso, fldr
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set fldr = fso.CreateFolder("C:\MyTest")
   msgbox "Created folder: " & fldr.Name
End Sub


'WRITING TEXT TO A TEXT FILE

	'Write
	'WriteLine
	'WriteBlankLines
	'Read
	'SkipLine
	'ReadLine
	'ReadAll

	'OpenAsTextFile
	'OpenAsTextStream
	'AtEndOfStream

Dim fso, f1, ts
Const ForWriting = 2
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile ("c:\test1.txt")
Set f1 = fso.GetFile("c:\test1.txt")
Set ts = f1.OpenAsTextStream(ForWriting, True)


dim FSO as scripting.filesystem object 'definition of rreferenced library set fso = new scripting.filesystemobject ' start a new instance of the file system object

end of sub routines its better practice to release your variables by setting them equal to nothing

dim variable as string set it = to path of folder destination

dim fil as scripting.file
    set fil = fso.getfile('file path here)
        'then check properties with intellisense
            fil.'property here


to get the beginning portion of the file PATH use environ("userprofile"), username just gives the name but profile give the first half of the file path

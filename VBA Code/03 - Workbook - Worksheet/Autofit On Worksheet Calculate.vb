'calculate event that runs every time the sheet calculates

Private Sub Worksheet_Calculate()
	Columns("A:F").AutoFit
End Sub

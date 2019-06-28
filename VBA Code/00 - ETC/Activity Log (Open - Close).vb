'have it auto assign live values upon workbook open and using environ to grab user name
' you can make a log that tracks accessing of a file
Private Sub Workbook_Open()
	Dim activityLog As Worksheet
		Set activityLog = ThisWorkbook.sheets("Activity Log")
	Dim logLastRow As Long
Application.ScreenUpdating = False
activityLog.Visible = True
    logLastRow = activityLog.Cells(Rows.Count, 1).End(xlUp).Row + 1
        activityLog.Cells(logLastRow, 1) = Date
        activityLog.Cells(logLastRow, 2) = Time
        activityLog.Cells(logLastRow, 3) = Environ("UserName")
        activityLog.Cells(logLastRow, 4) = "Open"
            Application.DisplayAlerts = False
                ThisWorkbook.Save
            Application.DisplayAlerts = True
activityLog.Visible = False
Application.ScreenUpdating = True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
	Dim activityLog As Worksheet
		Set activityLog = ThisWorkbook.sheets("Activity Log")
	Dim logLastRow As Long
Application.ScreenUpdating = False
activityLog.Visible = True
    logLastRow = activityLog.Cells(Rows.Count, 1).End(xlUp).Row + 1
        activityLog.Cells(logLastRow, 1) = Date
        activityLog.Cells(logLastRow, 2) = Tim
        activityLog.Cells(logLastRow, 3) = Environ("UserName")
        activityLog.Cells(logLastRow, 4) = "Close"
            Application.DisplayAlerts = False
                ThisWorkbook.Save
            Application.DisplayAlerts = True
activityLog.Visible = False
Application.ScreenUpdating = True
End Sub
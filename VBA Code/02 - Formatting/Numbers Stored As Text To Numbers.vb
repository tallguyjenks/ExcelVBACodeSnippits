'number conversion code from text to number
    With Selection
        Selection.NumberFormat = "General"
        .Value = .Value
    End With

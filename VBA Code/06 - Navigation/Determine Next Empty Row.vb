Dim nextRow as long
nextRow = application.worksheetfunction.countA([A:A])+1

'or


dim lastRow as long
lastRow = cells(rows.count,1).end(xlup)

'or


Dim lastRow
lastRow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).ROW

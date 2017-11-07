Sub SetMarkers()
'set up the data needed by excel and python to know where to begin writing next time
Dim wb As Workbook
Set wb = ThisWorkbook
Dim ws As Worksheet
Set ws = wb.Sheets("Main")
Dim lrow As Long

ws.Range("K1").Value = ws.Range("L1").Value
lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Range("L1").Value = lrow

End Sub

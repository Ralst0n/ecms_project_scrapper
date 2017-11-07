Sub RemoveDups()
Dim wb As Workbook
Set wb = ThisWorkbook
Dim ws As Worksheet
Set ws = wb.Sheets("Main")

ws.Cells.RemoveDuplicates Columns:=Array(1)




End Sub

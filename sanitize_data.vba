Sub SanitizeData()
Dim wb As Workbook
Set wb = ThisWorkbook
Dim ws As Worksheet
Set ws = wb.Sheets("Main")
Dim lrow As Long

lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'sanitize new data by removing leading and trailing white spaces
For Each c In ws.Range("A" + CStr((ws.Range("L1").Value) + 1), "G" + CStr(lrow))
    s = c.Value
        If Trim(Application.Clean(s)) <> s Then
            s = Trim(Application.Clean(s))
            c.Value = s
        End If

Next c


End Sub

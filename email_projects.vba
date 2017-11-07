Sub EmailProjects()
Dim wb As Workbook: Set wb = ThisWorkbook
Dim OutApp As Object
Dim OutMail As Object
Dim strbody As String
Dim ws As Worksheet: Set ws = wb.Sheets("Main")
Dim lrow As Long
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)
Dim email_array() As String
ReDim email_array(1 To 1) As String
Dim old_val As Integer: old_val = ws.Range("K1").Value
Dim new_val As Integer: new_val = ws.Range("L1").Value
Dim new_range As Range
If old_val <> new_val Then
Dim a As Range, b As Range

Set a = Range("A" & (old_val + 1), "H" & new_val)

For Each b In a.Rows
   ' MsgBox (b.Cells(1, 1) & ": " & b.Cells(1, 4) & " in " & b.Cells(1, 3) & " anticpated advance " & b.Cells(1, 6) & vbCrLf & b.Cells(1, 7))
    email_array(UBound(email_array)) = (b.Cells(1, 3) & " -- " & b.Cells(1, 4) & " (" & b.Cells(1, 1) & "):")
    ReDim Preserve email_array(1 To UBound(email_array) + 1) As String
    email_array(UBound(email_array)) = b.Cells(1, 7) & " ANTICIPATED ADV: " & b.Cells(1, 6)
    ReDim Preserve email_array(1 To UBound(email_array) + 1) As String
    email_array(UBound(email_array)) = b.Cells(1, 8)
    ReDim Preserve email_array(1 To UBound(email_array) + 1) As String
Next

  'i hate vba arrays must have one too many elements to be dynamic. deleting the last (blank) element
ReDim Preserve email_array(LBound(email_array) To UBound(email_array) - 1)
If (new_val - old_val) = 1 Then
    strbody = "<H2><B>The following project has been proposed on ECMS:</b></H2>"
Else
    strbody = "<H2><B>The following projects have been proposed on ECMS:</b></H2>"
End If
Index = 2
For Each Line In email_array
    Index = Index + 1
    If Index Mod 3 = 0 Then
        strbody = strbody + "<ul>"
        strbody = strbody + "<li><b><u>"
        strbody = strbody + Line
        strbody = strbody + "</b></u></li>"


    ElseIf Index Mod 3 = 1 Then
        strbody = strbody + "<p>"
        strbody = strbody + Line
        strbody = strbody + "</p>"


    Else
       strbody = strbody + "<p><i>"
       strbody = strbody + "Project Page: " + Line
       strbody = strbody + "</i></p>"
       strbody = strbody + "</ul>"
    End If

Next Line

With OutMail
        .To = "pajobs@prudenteng.com"
        .CC = ""
        .BCC = "rlawson@prudenteng.com"
        If new_val - old_val = 1 Then
        .Subject = "1 New proposed project: " & Format(Date, "mmmm dd, yyyy")
        Else
        .Subject = Str(new_val - old_val) & " New proposed projects: " & Format(Date, "mmmm dd, yyyy")
        End If
        .HTMLBody = strbody
        'Add attachments
        '.Attachments.Add ("C:\test.txt")
        .Send
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
  End If

'update persistence cells in workbook
ws.Range("K1") = ws.Range("L1")
ws.Range("F1") = Date


End Sub

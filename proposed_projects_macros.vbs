Dim args, objExcel

Set args = WScript.Arguments
Set objExcel = CreateObject("Excel.Application")

objExcel.Workbooks.Open("P:\010 - OFFICE\Office Projects\pruwork\ECMS Proposed Projects\Main\ECMS_proposed_projects.xlsm")
objExcel.Visible = False

objExcel.Run "SanitizeData"
objExcel.Run "RemoveDups"
objExcel.Run "SetMarkers"
objExcel.Run "EmailProjects"

objExcel.ActiveWorkbook.Save
objExcel.ActiveWorkbook.Close(0)
objExcel.Quit
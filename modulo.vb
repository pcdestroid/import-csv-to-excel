
Dim Url As String
Url = "" 'Insert google sheets url here (csv share link)
Dim excel: Set excel = CreateObject("Excel.Application")
excel.Application.DisplayAlerts = False
excel.Visible = False
excel.Workbooks.OpenText Url, 65001, , 1, , , , , , , True, ","
excel.ActiveWorkbook.SaveAs "C:\file.xlsx", 51
excel.ActiveWorkbook.Close
Set excel = Nothing


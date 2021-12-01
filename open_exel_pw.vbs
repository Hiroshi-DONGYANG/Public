Option Explicit
Dim objParam
Dim strFilename
Dim strPassword
Dim objExcel
Dim objOpendExcel
Set objParam = WScript.Arguments
strFilename = objParam(0)
strPassword = objParam(1)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objOpendExcel = objExcel.Workbooks.Open(strFilename,,,,strPassword)
Set objExcel = Nothing

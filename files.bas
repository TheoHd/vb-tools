Sub ImportDataFromClosedFile()
	Dim Path As String, File As String
	Path = "[enter_your_path]"
	File = "[enter_your_close_file].[xls/xlsx]"
	ThisWorkbook.Names.Add "range", _
	RefersTo:="='" & Path & "[" & File & "]Sheet1'!&A&1:$F$10"
	With Sheets("Sheet2")
		.[A1:F10] = "=range"
		.[A1:F10].Copy
		Sheets("Sheet1").Range("A1").PasteSpecial xlPasteValues
		.[A1:F10].Clear
	End With
End Sub

REM  *****  BASIC  *****

Option Compatible

Global CSVData(1) As String

Sub StartUp()
	If Not ThisComponent.SupportsService("com.sun.star.sheet.SpreadsheetDocument") Then
		Exit Sub
	End If
	Call GetDocumentProperty("CSV_Module/Selector", CSVData(1), ",")
	ThisComponent.calculateAll
End Sub

Sub ChangeSelector()
	Dim sInput As String
	sInput = InputBox("Change default CSV Selector","CSV Module",CSVData(1))
	If len(sInput) = 1 Then
		CSVData(1) = sInput
		Call SetDocumentProperty("CSV_Module/Selector", sInput)
		ThisComponent.calculateAll
	Else
		MsgBox("invalid CSV Selector, must be only one character", 0 + 48, "CSV Module"
	End If
End Sub

Sub SaveAsCSV()
	Dim FilePath As String
	Dim FilePointer As Integer
	Dim Row As Range
	Dim Filter (0 to 0, 0 to 1) As String
	Filter(0,0) = "Comma-Separated Values"
	Filter(0,1) = "*.csv"
	Call OpenSaveDialog("Save .csv file", Filter, FilePath)
	If FilePath = "" Then
		MsgBox("invalid file path", 0 + 48, "CSV Module")
		Exit Sub
	End If
	FilePointer = FreeFile()
	Open FilePath For Output As #FilePointer
	For Each Row In ThisComponent.CurrentController.getSelection().getDataArray()
		Print #FilePointer, CSVLINE(Row)
	Next Row
	Close #FilePointer
End Sub

Function CSVLINE(ParamArray Bereiche() As Variant) As String
	Dim Ergebnis As String
	Dim i As Integer
	Dim Zelle As Variant
	Ergebnis = ""
	For i = LBound(Bereiche) To UBound(Bereiche)
		If IsArray(Bereiche(i)) Then
			For Each Zelle in Bereiche(i)
				Ergebnis = Ergebnis & Zelle & CSVData(1)
			Next Zelle
		Else
			Ergebnis = Ergebnis & Bereiche(i) & CSVData(1)
		End If
	Next i
	CSVLINE = TRIMLAST(Ergebnis)
End Function

Function DATETOTEXT(Zelle, Optional FormatString As String = "dd.mm.yyyy") As String
	DATETOTEXT = Format(Zelle, FormatString)
End Function

Function TIMETOTEXT(Zelle, Optional FormatString As String = "dd.mm.yyyy HH:MM") As String
	TIMETOTEXT = Format(Zelle, FormatString)
End Function

Function TRIMLAST(Zelle As String, Optional Length As Integer = 1) As String
	TRIMLAST = Zelle
	If Len(Zelle) > 0 Then
		TRIMLAST = left(Zelle, Len(Zelle)-Length)
	End If
End Function

Function SCANUP(rowIndex As Long, colIndex As Long) As String
	On Error GoTo HandleError
	Dim oSheet As Object, oCell As Object
	oSheet = ThisComponent.CurrentController.ActiveSheet
	rowIndex = rowIndex - 1
	colIndex = colIndex - 1
	Do While rowIndex >= 0
		oCell = oSheet.getCellByPosition(colIndex, rowIndex)
		If oCell.Type <> com.sun.star.table.CellContentType.EMPTY Then
			If oCell.Type = com.sun.star.table.CellContentType.VALUE Then
				SCANUP = oCell.Value
			Else
				SCANUP = oCell.String
			End If
			Exit Function
		End If
		rowIndex = rowIndex - 1
	Loop
	SCANUP = ""
	Exit Function
HandleError:
	SCANUP = "#ERR"
End Function

Function SCANUP_REF(cellRef As String) As String
	On Error GoTo HandleError
	Dim oSheet As Object, oCell As Object
	oSheet = ThisComponent.GetCurrentController.ActiveSheet
	oCell = oSheet.getCellRangeByName(cellRef)
	SCANUP_REF = SCANUP(oCell.CellAddress.Row + 1, oCell.CellAddress.Column + 1)
	Exit Function
HandleError:
	SCANUP = "#ERR"
End Function


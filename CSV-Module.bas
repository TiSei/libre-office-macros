REM  *****  BASIC  *****

Option Compatible

Global CSVData(1) As String

Sub StartUp()
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

Function CSVLINE(ParamArray Bereiche() As Variant) As String
	Dim Ergebnis As String
	Dim i As Integer
	Dim Bereich As Variant
	Dim Zelle As Variant
	
	Ergebnis = ""
	If UBound(Bereiche) = 0 Then
		For Each Zelle In Bereiche(0)
			Ergebnis = Ergebnis & Zelle & CSVData(1)
		Next Zelle
	Else
		For i = LBound(Bereiche) To UBound(Bereiche)
			Ergebnis = Ergebnis & CSVLINE(Bereiche(i)) & CSVData(1)
		Next i
	End If
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


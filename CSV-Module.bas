REM  *****  BASIC  *****

Option Compatible

Global CSV_Selector As String

Sub StartUp()
	Dim oProps As Object
	oProps = ThisComponent.DocumentProperties.UserDefinedProperties
	If oProps.getPropertySetInfo().HasPropertyByName("CSV_Module/Selector") Then
		CSV_Selector = oProps.GetPropertyValue("CSV_Module/Selector")
	End If
	If CSV_Selector = "" Then
		CSV_Selector = ","
	End If
    ThisComponent.calculateAll
End Sub

Sub ChangeSelector()
	Dim sInput As String
	sInput = InputBox("Change default CSV Selector","CSV Module",Global_Selector)
	If Not len(sInput) = 1 Then
		MsgBox("invalid CSV Selector, must be only one character", 0 + 48, "CSV Module"
	Else
		CSV_Selector = sInput
		oProps = ThisComponent.DocumentProperties.UserDefinedProperties
		If oProps.getPropertySetInfo().hasPropertyByName("CSV_Module/Selector") Then
			aProps.SetPropertyValue("CSV_Module/Selector", sInput)
    	Else
    		oProps.AddProperty("CSV_Module/Selector", 256, sInput)
        End If
	End If
End Sub

Function CSVLINE(ParamArray Bereiche() As Variant) As String
	Dim Selector As String
	Dim Ergebnis As String
	Dim i As Integer
	Dim Bereich As Variant
	Dim Zelle As Variant
	
	Selector = CSV_Selector
	
	Ergebnis = ""
	If UBound(Bereiche) = 0 Then
		For Each Zelle In Bereiche(0)
			Ergebnis = Ergebnis & Zelle & Selector
		Next Zelle
	Else
		For i = LBound(Bereiche) To UBound(Bereiche)
			Ergebnis = Ergebnis & CSVLINE(Bereiche(i)) & Selector
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

Function TRIMLAST(Zelle) As String
	TRIMLAST = Zelle
	If Len(Zelle) > 0 Then
		TRIMLAST = left(Zelle, Len(Zelle)-1)
	End If
End Function


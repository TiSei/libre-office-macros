REM  *****  BASIC  *****

Option Compatible

Sub GetDocumentProperty(PropName As String, ByRef Pointer As Any, Optional Default As Any)
	Dim oProps As Object
	Pointer = Default
	oProps = ThisComponent.DocumentProperties.UserDefinedProperties
	If oProps.getPropertySetInfo().HasPropertyByName(PropName) Then
		Pointer = oProps.GetPropertyValue(PropName)
	End If
End Sub

Sub SetDocumentProperty(PropName As String, ByVal Value As Any, Optional Flags As Integer = 256)
	Dim oProps As Object
	oProps = ThisComponent.DocumentProperties.UserDefinedProperties
	If oProps.getPropertySetInfo().hasPropertyByName(PropName) Then
		oProps.SetPropertyValue(PropName, Value)
	Else
		oProps.AddProperty(PropName, Flags, Value)
	End If
End Sub

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

Sub OpenSaveDialog(Title As String, Filters As Object, ByRef Pointer As String, Optional Mode As Integer = 10)
	Dim i As Integer
	Dim oFielDialog As Object
	oFileDialog = com.sun.star.ui.dialogs.FilePicker.createWithMode(Mode)
	With oFileDialog
		.setTitle(Title)
		For i = LBound(Filters) to UBound(Filters)
			.appendFilter(Filters(i,0),Filters(i,1))
		Next i
		.appendFilter("all files","*.*")
		If .Execute() <> 1 Then Exit Sub
		Pointer = .SelectedFiles(0)
	End With
End Sub

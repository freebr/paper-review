<%
Const msoTextBox=17
Const wdExportFormatPDF=17
Const wdGoToPage=1
Const wdGoToAbsolute=1
Const wdStatisticPages=2
' 论文相关表格生成器类
Class PaperFormWriter
	Dim m_info,m_err_desc
	Private Sub Class_Initialize()
		Set m_info=CreateDictionary()
	End Sub
	Private Sub Class_Terminate()
		Set m_info=Nothing
	End Sub
	Property Get ErrorDesc()
		ErrorDesc=m_err_desc
	End Property
	Function setField(key,value)
		m_info.Add key,value
		setField=1
	End Function
	Function setArray(prop_name,value)
		If IsEmpty(value) Then
			setField prop_name,Array()
			Exit Function
		End If
		Dim arr:arr=value
		If Not IsArray(arr) Then
			ReDim arr(value.Count-1)
			Dim i
			For i=0 To UBound(arr)
				arr(i)=value(i+1)
			Next
		End If
		setField prop_name,arr
		setArray=1
	End Function
	Function removeField(key)
		m_info.Remove key
		removeField=1
	End Function
	Function generateTable(filepath,template_name)
		' 生成论文相关表格
		Dim app,doc,template_file
		'On Error Resume Next
		template_file=Server.MapPath("template/"&template_name&".doc")
		Set app=Server.CreateObject("Word.Application")
		Set doc=app.Documents.Add(template_file)
		app.Run "fillInTable",m_info,filepath
		doc.Close(False)
		Set doc=Nothing
		app.Quit()
		Set app=Nothing
		m_err_desc=Err.Description
		If Err.Number=0 Then
			generateTable=1
		Else
			generateTable=0
		End If
	End Function
End Class
%>
<%
Const msoTextBox=17
Const wdExportFormatPDF=17
Const wdGoToPage=1
Const wdGoToAbsolute=1
Const wdStatisticPages=2
' 论文相关表格生成器类
Class TableGen
	Dim m_info,m_err_desc
	Private Sub Class_Initialize()
		Set m_info=Server.CreateObject("Scripting.Dictionary")
	End Sub
	Private Sub Class_Terminate()
		Set m_info=Nothing
	End Sub
	Public Property Get ErrorDesc()
		ErrorDesc=m_err_desc
	End Property
	Public Function addInfo(key,value)
		m_info.Add key,value
		addInfo=1
	End Function
	Public Function removeInfo(key)
		m_info.Remove key
		removeInfo=1
	End Function
	Public Function generateTable(filepath,template_name)
		' 生成论文相关表格
		Dim app,doc,template_file
		'On Error Resume Next
		template_file=Server.MapPath("template/"&template_name&".doc")
		Set app=Server.CreateObject("Word.Application")
		Set doc=app.Documents.Add(template_file)
		app.Run "fillInTable",m_info,filepath
		doc.Close(False)
		Set range=Nothing
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
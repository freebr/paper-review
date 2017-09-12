<%
' 预答辩意见书生成器类
Class PrereplyGen
	Dim m_author,m_enteryear,m_spec,m_thesis_subject
	Dim m_err_desc
	Public Property Let Author(val)
		m_author=val
	End Property
	Public Property Let EnterYear(val)
		m_enteryear=val
	End Property
	Public Property Let Spec(val)
		m_spec=val
	End Property
	Public Property Let Subject(val)
		m_thesis_subject=val
	End Property
	Public Property Get ErrorDesc()
		ErrorDesc=m_err_desc
	End Property
	Public Function generateTable(filepath)
		' 生成预答辩意见书
		Dim template_file
		Dim wd,doc,range
		'On Error Resume Next
		template_file=Server.MapPath("export")&"\ydbyjs.dot"
		Set wd=Server.CreateObject("Word.Application.12")
		Set doc=wd.Documents.Add(template_file)
		Set range=doc.Range()
		range.Select()
		range.Find.Execute "$author",True,,,,,,1,,toWordString(m_author),True
		range.Find.Execute "$enter_year",True,,,,,,1,,toWordString(m_enteryear),True
		range.Find.Execute "$spec",True,,,,,,1,,toWordString(m_spec),True
		range.Find.Execute "$thesis_subject",True,,,,,,1,,toWordString(m_thesis_subject),True
		doc.BuiltInDocumentProperties(3)="华南理工大学工商管理学院"
		doc.SaveAs filepath
		doc.Close(False)
		wd.Quit()
		Set range=Nothing
		Set doc=Nothing
		Set wd=Nothing
		m_err_desc=Err.Description
		If Err.Number=0 Then
			generateTable=1
		Else
			generateTable=0
		End If
	End Function
	Private Function toWordString(s)
		If IsNull(s) Then Exit Function
		Dim ret
		ret=s
		ret=Replace(ret,"^","^^")
		ret=Replace(ret,vbNewLine,"^p")
		ret=Replace(ret,vbCr,"^p")
		ret=Replace(ret,vbLf,"^p")
		toWordString=ret
	End Function
End Class
%>
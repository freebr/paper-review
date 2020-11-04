<%
' 送审申请表生成器类
Class ReviewApplicationFormWriter
	Dim m_author,m_stuno,m_tutorinfo,m_spec,m_date,m_thesis_subject,m_comment,m_reproduct_ratio
	Dim m_err_desc
	Public Property Let Author(val)
		If IsNull(val) Then val=""
		m_author=val
	End Property
	Public Property Let StuNo(val)
		If IsNull(val) Then val=""
		m_stuno=val
	End Property
	Public Property Let TutorInfo(val)
		If IsNull(val) Then val=""
		m_tutorinfo=val
	End Property
	Public Property Let Spec(val)
		If IsNull(val) Then val=""
		m_spec=val
	End Property
	Public Property Let Date(val)
		If IsNull(val) Then val=""
		m_date=val
	End Property
	Public Property Let Subject(val)
		If IsNull(val) Then val=""
		m_thesis_subject=val
	End Property
	Public Property Let Comment(val)
		If IsNull(val) Then val=""
		m_comment=val
	End Property
	Public Property Let ReproductRatio(val)
		If IsNull(val) Then val=""
		m_reproduct_ratio=val
	End Property
	Public Property Get ErrorDesc()
		ErrorDesc=m_err_desc
	End Property
	Public Function generateApp(filepath)
		' 生成送审申请表
		Dim template_file
		Dim wd,doc,range,text,i
		'On Error Resume Next
		template_file=Server.MapPath(basePath()&"tutor/export/sssqb.dot")
		Set wd=Server.CreateObject("Word.Application")
		Set doc=wd.Documents.Add(template_file)
		Dim val,bValid
		For i=1 To doc.Shapes.Count
			bValid=True
			Set range=doc.Shapes(i).TextFrame.TextRange
			text=Trim(Left(range.Text,Len(range.Text)-1))
			Select Case text
			Case "$author":val=m_author
			Case "$stuno":val=m_stuno
			Case "$tutorinfo":val=m_tutorinfo
			Case "$spec":val=m_spec
			Case "$date":val=m_date
			Case Else:bValid=False
			End Select
			If bValid Then
				range.Select()
				range.Delete()
				range.InsertAfter(val)
			End If
		Next
		Set range=doc.Range()
		range.Select()
		range.Find.Execute "$thesis_subject",True,,,,,,1,,toWordString(m_thesis_subject),True
		If range.Find.Execute("$comment",True,,,,,,1) Then
			range.Text=m_comment
		End If
		range.Find.Execute "$reproduct_ratio",True,,,,,,1,,toWordString(m_reproduct_ratio),True
		doc.BuiltInDocumentProperties(3)="华南理工大学工商管理学院"
		doc.SaveAs filepath,16
		doc.Close(False)
		Set range=Nothing
		Set doc=Nothing
		wd.Quit()
		Set wd=Nothing
		m_err_desc=Err.Description
		If Err.Number=0 Then
			generateApp=1
		Else
			generateApp=0
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
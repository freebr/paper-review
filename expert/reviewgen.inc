<%
Const msoTextBox=17
Const wdExportFormatPDF=17
Const wdGoToPage=1
Const wdGoToAbsolute=1
Const wdStatisticPages=2
' 评阅书生成器类
Class ReviewGen
	Dim m_author,m_tutorinfo,m_spec,m_researchway,m_date,m_thesis_subject
	Dim m_expert_name,m_expert_pro_duty,m_expert_expertise,m_expert_workplace,m_expert_address,m_expert_mailcode
	Dim m_expert_tel1,m_expert_tel2,m_expert_master_level,m_eval_text,m_correlation_type,m_reviewresult
	Dim m_scores,m_scoreparts,m_totalscore,m_reviewlevel,m_thesistype
	Dim m_err_desc
	Public Property Let Author(val)
		m_author=val
	End Property
	Public Property Let TutorInfo(val)
		m_tutorinfo=val
	End Property
	Public Property Let Spec(val)
		m_spec=val
	End Property
	Public Property Let ResearchWay(val)
		m_researchway=val
	End Property
	Public Property Let Date(val)
		m_date=val
	End Property
	Public Property Let Subject(val)
		m_thesis_subject=val
	End Property
	Public Property Let ExpertName(val)
		m_expert_name=val
	End Property
	Public Property Let ExpertProDuty(val)
		m_expert_pro_duty=val
	End Property
	Public Property Let ExpertExpertise(val)
		m_expert_expertise=val
	End Property
	Public Property Let ExpertWorkplace(val)
		m_expert_workplace=val
	End Property
	Public Property Let ExpertAddress(val)
		m_expert_address=val
	End Property
	Public Property Let ExpertMailcode(val)
		m_expert_mailcode=val
	End Property
	Public Property Let ExpertTel1(val)
		m_expert_tel1=val
	End Property
	Public Property Let ExpertTel2(val)
		m_expert_tel2=val
	End Property
	Public Property Let ExpertMasterLevel(val)
		m_expert_master_level=val
	End Property
	Public Property Let EvalText(val)
		m_eval_text=val
	End Property
	Public Property Let CorrelationType(val)
		m_correlation_type=val
	End Property
	Public Property Let ReviewResult(val)
		m_reviewresult=val
	End Property
	Public Property Let ReviewLevel(val)
		m_reviewlevel=val
	End Property
	Public Property Let ThesisType(val)
		m_thesistype=val
	End Property
	Public Property Let Scores(s)
		m_scores=Split(s,",")
	End Property
	Public Property Let ScoreParts(s)
		m_scoreparts=Split(s,",")
	End Property
	Public Property Let TotalScore(val)
		m_totalscore=val
	End Property
	Public Property Get ErrorDesc()
		ErrorDesc=m_err_desc
	End Property
	Public Function generateReview(filepath,filepath2,filepath3,template_file,reviewfile_type)
		' 生成评阅书
		Dim wd,doc,range,range2,text
		Dim i,l,ftsize
		'On Error Resume Next
		Set wd=Server.CreateObject("Word.Application")
		Set doc=wd.Documents.Add(template_file)
		Dim val,bValid
		For i=1 To doc.Shapes.Count
			If doc.Shapes(i).Type=msoTextBox Then
				bValid=True
				Set range=doc.Shapes(i).TextFrame.TextRange
				text=Trim(Left(range.Text,Len(range.Text)-1))
				Select Case text
				Case "$expert_name":val=m_expert_name
				Case "$expert_pro_duty":val=m_expert_pro_duty
				Case "$expert_expertise":val=m_expert_expertise
				Case "$expert_workplace":val=m_expert_workplace
				Case "$expert_address":val=m_expert_address
				Case "$expert_mailcode":val=m_expert_mailcode
				Case "$expert_tel1":val=m_expert_tel1
				Case "$expert_tel2":val=m_expert_tel2
				Case "$expert_master_level":val=String(m_expert_master_level,vbTab)&"√"
				Case "$correlation_type":val=String(m_correlation_type,vbTab)&"√"
				Case "$review_result"
					If reviewfile_type=1 Then
						val=String(m_reviewresult,vbTab)&"√"
					Else
						val=String((m_reviewresult-1)\2,vbNewLine)&String((m_reviewresult-1)Mod 2+1,vbTab)&"√"
					End If
				Case "$review_level":val=String(m_reviewlevel,vbTab)&"√"
				Case "$thesis_type"
					If m_thesistype>=12 Then
						val=String(m_thesistype-11,vbTab)&"√"
					Else
						val=""
					End If
				Case Else:bValid=False
				End Select
				If bValid Then
					range.Select()
					range.Find.Execute text,True,,,,,,1,,toWordString(val),True
				End If
			End If
		Next
		Set range=doc.Range()
		range.Select()
		Do
			bFind=range.Find.Execute("$author",True,,,,,,1,,toWordString(m_author),True)
		Loop While bFind
		
		' 调整字体大小
		l=Len(m_thesis_subject)
		If l<=17 Then
			ftsize=16
		ElseIf l<=20 Then
			ftsize=14
		ElseIf l<=23 Then
			ftsize=12
		Else
			ftsize=10.5
		End If
		Set range2=doc.Tables(1).Range
		If range2.Find.Execute("$thesis_subject",True) Then
			range2.Font.Size=ftsize
		End If
		Do
			bFind=range.Find.Execute("$thesis_subject",True,,,,,,1,,toWordString(m_thesis_subject),True)
		Loop While bFind
		range.Find.Execute "$tutorinfo",True,,,,,,1,,toWordString(m_tutorinfo),True
		range.Find.Execute "$spec",True,,,,,,1,,toWordString(m_spec),True
		range.Find.Execute "$researchway",True,,,,,,1,,toWordString(m_researchway),True
		range.Find.Execute "$submit_review_date",True,,,,,,1,,toWordString(m_date),True
		If range.Find.Execute("$eval_text",True,,,,,,1) Then
			range.Text=m_eval_text
		End If
		If reviewfile_type=2 Then
			' 对ME/MBA评阅书进行处理
			i=doc.ComputeStatistics(wdStatisticPages)
			j=doc.Tables(3).Range.End
			Set r1=doc.GoTo(wdGoToPage,wdGoToAbsolute,i-1)
			Set r2=doc.GoTo(wdGoToPage,wdGoToAbsolute,i)
			If Len(doc.Range(r1.Start,r2.Start))=2 Then
				' 删除多余的换页符
				doc.Range(j,j).Delete()
			End If
			Set r1=Nothing
			Set r2=Nothing
			For i=0 To UBound(m_scores)
				range.Find.Execute "$<score"&(i+1)&">",True,,True,,,,1,,toWordString(Int(m_scores(i))),True
			Next
			If IsArray(m_scoreparts) Then
				For i=0 To UBound(m_scoreparts)
					range.Find.Execute "$<scorep"&(i+1)&">",True,,True,,,,1,,toWordString(Int(m_scoreparts(i))),True
				Next
			End If
		End If
		range.Find.Execute "$<score>",True,,True,,,,1,,toWordString(m_totalscore),True
		' 生成完整版评阅书
		doc.ExportAsFixedFormat filepath,wdExportFormatPDF,,1
		' 生成不含学生信息的评阅书
		doc.Tables(1).Delete()
		If reviewfile_type=1 Then
			doc.Range.Paragraphs(4).Range.Delete()
		Else
			doc.Range.Paragraphs(3).Range.Delete()
		End If
		doc.Tables(2).Cell(1,3).Delete()
		doc.Tables(2).Cell(1,2).Delete()
    doc.Tables(2).Cell(1,1).Range.Text="论文题目"
    doc.Tables(2).Cell(1,2).SetWidth 425.7,1
		doc.ExportAsFixedFormat filepath2,wdExportFormatPDF,,1
		' 生成不含专家信息的评阅书
		doc.Undo 6
		doc.Tables(2).Range.Delete()
		doc.ExportAsFixedFormat filepath3,wdExportFormatPDF,,1
		doc.Close(False)
		wd.Quit()
		Set range2=Nothing
		Set range=Nothing
		Set doc=Nothing
		Set wd=Nothing
		m_err_desc=Err.Description
		If Err.Number=0 Then
			generateReview=1
		Else
			generateReview=0
		End If
	End Function
	Private Function toWordString(s)
		If IsNull(s) Then Exit Function
		Dim ret
		ret=s
		ret=Replace(ret,"^","^^")
		ret=Replace(ret,vbNewLine,"^p")
		ret=Replace(ret,vbCr,"^p")
		ret=Replace(ret,vbLf,"^l")
		ret=Replace(ret,vbTab,"^t")
		toWordString=ret
	End Function
End Class
%>
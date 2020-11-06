<%
Const msoTextBox=17
Const wdExportFormatPDF=17
Const wdGoToPage=1
Const wdGoToAbsolute=1
Const wdStatisticPages=2
' 评阅书生成器类
Class ReviewDocumentWriter
	Dim m_author,m_stu_no,m_tutor_info,m_speciality,m_researchway,m_date,m_review_pattern,m_subject
	Dim m_expert_name,m_expert_pro_duty,m_expert_expertise,m_expert_workplace,m_expert_address,m_expert_mailcode
	Dim m_expert_tel1,m_expert_tel2,m_expert_master_level,m_comment,m_suggestion,m_correlation_level,m_review_result
	Dim m_scores,m_scoreparts,m_total_score,m_review_level,m_paper_type_name
	Dim m_err_desc
	Public Property Let Author(val)
		m_author=val
	End Property
	Public Property Let StuNo(val)
		m_stu_no=val
	End Property
	Public Property Let TutorInfo(val)
		m_tutor_info=val
	End Property
	Public Property Let Spec(val)
		m_speciality=val
	End Property
	Public Property Let ResearchWay(val)
		m_researchway=val
	End Property
	Public Property Let Date(val)
		m_date=val
	End Property
	Public Property Let ReviewPattern(val)
		m_review_pattern=val
	End Property
	Public Property Let Subject(val)
		m_subject=val
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
	Public Property Let Comment(val)
		m_comment=val
	End Property
	Public Property Let Suggestion(val)
		m_suggestion=val
	End Property
	Public Property Let CorrelationLevel(val)
		m_correlation_level=val
	End Property
	Public Property Let ReviewResult(val)
		m_review_result=val
	End Property
	Public Property Let ReviewLevel(val)
		m_review_level=val
	End Property
	Public Property Let PaperTypeName(val)
		m_paper_type_name=val
	End Property
	Public Property Let Scores(s)
		m_scores=s
	End Property
	Public Property Let ScoreParts(s)
		m_scoreparts=s
	End Property
	Public Property Let TotalScore(val)
		m_total_score=val
	End Property
	Public Property Get ErrorDesc()
		ErrorDesc=m_err_desc
	End Property

	Public Function exportReviewDocument(filepath,filepath2,filepath3,template_file,stu_type,wd)
		' 生成评阅书
		Dim doc,info
		If IsNull(wd) Then Set wd = Server.CreateObject("Word.Application")
		Set doc=wd.Documents.Add(template_file)
		On Error Resume Next
		Set info=CreateDictionary()
		info("Author") = m_author
		info("TutorInfo") = m_tutor_info
		info("Subject") = m_subject
		info("StuNo") = m_stu_no
		info("Speciality")=m_speciality
		info("Researchway")=m_researchway
		info("ReviewPattern") = m_review_pattern
		info("SubmitReviewDate") = m_date
		info("ExpertName") = m_expert_name
		info("ExpertProDuty") = m_expert_pro_duty
		info("ExpertExpertise") = m_expert_expertise
		info("ExpertWorkplace") = m_expert_workplace
		info("ExpertAddress") = m_expert_address
		info("ExpertMailcode") = m_expert_mailcode
		info("ExpertTel1") = m_expert_tel1
		info("ExpertTel2") = m_expert_tel2
		info("ExpertMasterLevel") = m_expert_master_level
		info("CorrelationLevel") = m_correlation_level
		info("PaperTypeName") = m_paper_type_name
		info("Scores") = m_scores
		info("ScoreParts") = m_scoreparts
		info("TotalScore") = m_total_score
		info("Comment") = m_comment
		info("Suggestion") = m_suggestion
		info("ReviewLevel") = m_review_level
		info("ReviewResult") = m_review_result
		info("FilePath") = filepath
		info("FilePath2") = filepath2
		info("FilePath3") = filepath3
		wd.Run "exportReviewDocument", info
		If Err.Number=438 Then	' 未设置宏
			On Error GoTo 0
			exportReviewDocumentDefault doc,stu_type,info,filepath,filepath2,filepath3
		Else
			On Error GoTo 0
		End If
		doc.Close(False)
		Set doc=Nothing
		If IsNull(wd) Then
			wd.Quit()
			Set wd=Nothing
		End If
		m_err_desc=Err.Description
		If Err.Number=0 Then
			exportReviewDocument=1
		Else
			exportReviewDocument=0
		End If
	End Function

	Function exportReviewDocumentDefault(doc,stu_type,info,filepath,filepath2,filepath3)
		Dim range,range2,text
		Dim i,l,font_size
		Dim val,bValid,bFind
		For i=1 To doc.Shapes.Count
			If doc.Shapes(i).Type=msoTextBox Then
				bValid=True
				Set range=doc.Shapes(i).TextFrame.TextRange
				text=Trim(Left(range.Text,Len(range.Text)-1))
				Select Case text
				Case "$expert_name":val=info("ExpertName")
				Case "$expert_pro_duty":val=info("ExpertProDuty")
				Case "$expert_expertise":val=info("ExpertExpertise")
				Case "$expert_workplace":val=info("ExpertWorkplace")
				Case "$expert_address":val=info("ExpertAddress")
				Case "$expert_mailcode":val=info("ExpertMailcode")
				Case "$expert_tel1":val=info("ExpertTel1")
				Case "$expert_tel2":val=info("ExpertTel2")
				Case "$expert_master_level":val=String(info("ExpertMasterLevel"),vbTab)&"√"
				Case "$correlation_level":val=String(info("CorrelationLevel"),vbTab)&"√"
				Case "$review_result"
					If stu_type=5 Then
						val=String((info("ReviewResult")-1)\2,vbNewLine)&String((info("ReviewResult")-1)Mod 2+1,vbTab)&"√"
					Else
						val=String(info("ReviewResult"),vbTab)&"√"
					End If
				Case "$review_level":val=String(info("ReviewLevel"),vbTab)&"√"
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
			bFind=range.Find.Execute("$author",True,,,,,,1,,toWordString(info("Author")),True)
		Loop While bFind
		
		' 调整字体大小
		l=Len(info("Subject"))
		If l<=17 Then
			font_size=16
		ElseIf l<=20 Then
			font_size=14
		ElseIf l<=23 Then
			font_size=12
		Else
			font_size=10.5
		End If
		Set range2=doc.Tables(1).Range
		If range2.Find.Execute("$thesis_subject",True) Then
			range2.Font.Size=font_size
		End If
		Do
			bFind=range.Find.Execute("$thesis_subject",True,,,,,,1,,toWordString(info("Subject")),True)
		Loop While bFind
		range.Find.Execute "$tutorinfo",True,,,,,,1,,toWordString(info("TutorInfo")),True
		range.Find.Execute "$spec",True,,,,,,1,,toWordString(info("Speciality")),True
		range.Find.Execute "$researchway",True,,,,,,1,,toWordString(info("Researchway")),True
		range.Find.Execute "$submit_review_date",True,,,,,,1,,toWordString(info("SubmitReviewDate")),True
		If range.Find.Execute("$eval_text",True,,,,,,1) Then
			range.Text=info("Comment")
		End If
		If Len(info("Scores")) Then
			Dim scores:scores=Split(info("Scores"),",")
			' 对ME评阅书进行处理
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
			For i=0 To UBound(scores)
				range.Find.Execute "$score"&(i+1),True,,True,,,,1,,toWordString(Int(scores(i))),True
			Next
		End If
		If Len(info("ScoreParts")) Then
			Dim scoreparts:scoreparts=Split(info("ScoreParts"),",")
			For i=0 To UBound(scoreparts)
				range.Find.Execute "$scorep"&(i+1),True,,True,,,,1,,toWordString(Int(scoreparts(i))),True
			Next
		End If
		range.Find.Execute "$total_score",True,,True,,,,1,,toWordString(info("TotalScore")),True
		' 生成完整版评阅书
		doc.ExportAsFixedFormat filepath,wdExportFormatPDF,,1
		' 生成不含学生信息的评阅书
		doc.Tables(1).Delete()
		If stu_type=5 Then
			doc.Range.Paragraphs(3).Range.Delete()
		Else
			doc.Range.Paragraphs(4).Range.Delete()
		End If
		doc.Tables(2).Cell(1,3).Delete()
		doc.Tables(2).Cell(1,2).Delete()
		doc.Tables(2).Cell(1,1).Range.Text="论文题目"
		doc.Tables(2).Cell(1,2).SetWidth 412.65,1
		doc.ExportAsFixedFormat filepath2,wdExportFormatPDF,,1
		' 生成不含专家信息的评阅书
		doc.Undo 6
		doc.Tables(2).Range.Delete()
		doc.ExportAsFixedFormat filepath3,wdExportFormatPDF,,1
		Set range2=Nothing
		Set range=Nothing
		exportReviewDocumentDefault=1
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
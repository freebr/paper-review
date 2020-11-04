<!--#include file="exportReviewDocumentDefault.inc"-->
<%
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
	Public Function exportReviewDocument(filepath,filepath2,filepath3,template_file,stu_type)
		' 生成评阅书
		Dim wd,doc,info
		Set wd=Server.CreateObject("Word.Application")
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
		If Err.Number=-2147352573 Then	' 未设置宏
			On Error GoTo 0
			exportReviewDocumentDefault doc,stu_type,filepath,filepath2,filepath3
		Else
			On Error GoTo 0
		End If
		doc.Close(False)
		Set doc=Nothing
		wd.Quit()
		Set wd=Nothing
		m_err_desc=Err.Description
		If Err.Number=0 Then
			exportReviewDocument=1
		Else
			exportReviewDocument=0
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
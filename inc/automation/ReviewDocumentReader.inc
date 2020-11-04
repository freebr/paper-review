<%
' 评阅书读取器类
Class ReviewDocumentReader
	Dim m_stu_no
	Dim m_expert_name,m_expert_pro_duty,m_expert_expertise,m_expert_workplace
	Dim m_expert_address,m_expert_tel,m_expert_master_level
	Dim m_comment,m_suggestion,m_correlation_level,m_reviewresult,m_review_time,m_review_pattern
	Dim m_scores,m_scoreparts,m_totalscore,m_reviewlevel
	Dim m_err_desc
	Public Property Get StuNo()
		StuNo=m_stu_no
	End Property
	Public Property Get ExpertName()
		ExpertName=m_expert_name
	End Property
	Public Property Get ExpertMasterLevel()
		ExpertMasterLevel=m_expert_master_level
	End Property
	Public Property Get ExpertProDuty()
		ExpertProDuty=m_expert_pro_duty
	End Property
	Public Property Get ExpertExpertise()
		ExpertExpertise=m_expert_expertise
	End Property
	Public Property Get ExpertWorkplace()
		ExpertWorkplace=m_expert_workplace
	End Property
	Public Property Get ExpertAddress()
		ExpertAddress=m_expert_address
	End Property
	Public Property Get ExpertTel()
		ExpertTel=m_expert_tel
	End Property
	Public Property Get Comment()
		Comment=m_comment
	End Property
	Public Property Get Suggestion()
		Suggestion=m_suggestion
	End Property
	Public Property Get ReviewPattern()
		ReviewPattern=m_review_pattern
	End Property
	Public Property Get CorrelationLevel()
		CorrelationLevel=m_correlation_level
	End Property
	Public Property Get ReviewResult()
		ReviewResult=m_reviewresult
	End Property
	Public Property Get ReviewLevel()
		ReviewLevel=m_reviewlevel
	End Property
	Public Property Get ReviewTime()
		ReviewTime=m_review_time
	End Property
	Public Property Get Scores()
		Scores=m_scores
	End Property
	Public Property Get ScoreParts()
		ScoreParts=m_scoreparts
	End Property
	Public Property Get TotalScore()
		TotalScore=m_totalscore
	End Property
	Public Property Get ErrorDesc()
		ErrorDesc=m_err_desc
	End Property
	Public Function extractInfoFromReviewDocument(filepath)
		' 读取评阅书 Word 文档中的评阅信息
		Dim wd, doc
		Set wd = Server.CreateObject("Word.Application")
		Set doc = wd.Documents.Add(filepath)
		' On Error Resume Next
		' 从文件名获取专家姓名
		m_expert_name = Split(Mid(filepath, InStrRev(filepath, "\") + 1), "_")(0)
		m_scores = Null
		m_scoreparts = Null
		' 获取学生类型
		Dim stu_type: stu_type = compactString(doc.Tables(2).Cell(4, 2).Range.Text, True)
		Select Case stu_type
			Case "工商管理硕士"
				If doc.Tables.Count >= 13 Then	' MBA
					extractMBAInfo doc
				Else	' EMBA
					extractEMBAInfo doc
				End If
			Case "会计硕士"	' MPAcc
				extractMPAccInfo doc
			Case "高级管理人员工商管理硕士"	' EMBA
				extractEMBAInfo doc
			Case Else	' MEM
				If doc.Range.Find.Execute("工程硕士学位论文评阅书") Then
					extractMEMInfo doc
				Else
					Err.Raise "无法识别评阅书对应的学生类型。"
				End If
		End Select
		doc.Close(False)
		Set doc=Nothing
		wd.Quit()
		Set wd=Nothing
		m_err_desc=Err.Description
		If Err.Number=0 Then
			extractInfoFromReviewDocument = 1
		Else
			extractInfoFromReviewDocument = 0
		End If
	End Function
	' 提取 MBA 类型评阅书信息
	Private Function extractMBAInfo(doc)
		Dim tmp
		' 送审时间
		m_review_time = DateValue(compactString(doc.Tables(2).Cell(9, 2).Range.Text, True))
		' 送审方式
		m_review_pattern = compactString(doc.Tables(2).Cell(8, 2).Range.Text, True)
		' 学号
		m_stu_no = compactString(doc.Tables(2).Cell(7, 2).Range.Text, True)
		' 评语
		Dim rg: Set rg = doc.Tables(12).Cell(3, 1).Range
		Dim i
		For i = 2 To rg.Sentences.Count
			tmp = tmp & rg.Sentences(i).Text
		Next
		m_comment = compactString(tmp, False)
		' 论文存在的不足之处和建议等
		tmp = ""
		Set rg = doc.Tables(12).Cell(4, 1).Range
		For i = 2 To rg.Sentences.Count
			tmp = tmp & rg.Sentences(i).Text
		Next
		m_suggestion = compactString(tmp, False)
		' 评分
		tmp = ""
		Dim tbl: Set tbl = doc.Tables(10)
		For i = 2 To tbl.Rows.Count - 1
			If i > 2 Then tmp = tmp & ","
            tmp = tmp & getScore(tbl.Cell(i, 3).Range.Text)
		Next
		' 指标得分
		m_scores = Split(tmp, ",")
		' 加权总分
		m_totalscore = getScore(tbl.Cell(tbl.Rows.Count, 2).Range.Text)
		' 评阅专家信息
		' 专业技术职务
		m_expert_pro_duty = compactString(doc.Tables(4).Cell(1, 4).Range.Text, True)
		' 学科专长
		m_expert_expertise = compactString(doc.Tables(5).Cell(1, 2).Range.Text, True)
		' 单位名称（含院系）
		m_expert_workplace = compactString(doc.Tables(6).Cell(1, 2).Range.Text, True)
		' 通信地址
		m_expert_address = compactString(doc.Tables(7).Cell(1, 2).Range.Text, True)
		' 联系电话
		m_expert_tel = compactString(doc.Tables(7).Cell(1, 4).Range.Text, True)
		' 对论文涉及内容的熟悉程度
		Dim arrMasterLevelText: arrMasterLevelText = Array("", "很熟悉(√)", "熟悉(√)", "一般(√)")
		For i = 1 To UBound(arrMasterLevelText)
			If doc.Tables(7).Cell(2, 1).Range.Find.Execute(arrMasterLevelText(i)) Then
				m_expert_master_level = i
				Exit For
			End If
		Next
		' 对学位论文的总体评价
		For i = 1 To 4
			If doc.Tables(13).Cell(2, i + 1).Range.Find.Execute("√") Then
				m_reviewlevel = i
				Exit For
			End If
		Next
		' 是否同意申请论文答辩
		Set rg = doc.Tables(13).Cell(3, 2).Range
		Dim j: j = 1
		For i = 1 To rg.Paragraphs.Count
			tmp = compactString(rg.Paragraphs(i).Range.Text, True)
			If Len(tmp) > 1 Then
				If rg.Paragraphs(i).Range.Find.Execute("√") Then
					m_reviewresult = j
					Exit For
				End If
				j = j + 1
			End If
		Next
		' 学位论文内容与申请学位专业的相关性
		m_correlation_level = 1
		extractMBAInfo = 1
	End Function
	' 提取 MEM 类型评阅书信息
	Private Function extractMEMInfo(doc)
		Dim tmp
		' 送审时间
		m_review_time = DateValue(compactString(doc.Tables(2).Cell(8, 2).Range.Text, True))
		' 送审方式
		m_review_pattern = compactString(doc.Tables(2).Cell(8, 2).Range.Text, True)
		' 学号
		m_stu_no = compactString(doc.Tables(9).Cell(1, 4).Range.Text, True)
		' 评语
		Dim rg: Set rg = doc.Tables(9).Cell(3, 1).Range
		Dim i
		For i = 2 To rg.Sentences.Count
			tmp = tmp & rg.Sentences(i).Text
		Next
		m_comment = compactString(tmp, False)
		' 评分
		tmp = ""
		Dim tbl: Set tbl = doc.Tables(10)
		For i = 2 To tbl.Rows.Count - 1
			If i > 2 Then tmp = tmp & ","
			tmp = tmp & getScore(tbl.Cell(i, 5).Range.Text)
		Next
		' 指标得分
		m_scores = Split(tmp, ",")
		' 加权总分
		m_totalscore = getScore(tbl.Cell(tbl.Rows.Count, 2).Range.Text)
		' 评阅专家信息
		' 专业技术职务
		m_expert_pro_duty = compactString(doc.Tables(4).Cell(1, 4).Range.Text, True)
		' 学科专长
		m_expert_expertise = compactString(doc.Tables(5).Cell(1, 2).Range.Text, True)
		' 单位名称（含院系）
		m_expert_workplace = compactString(doc.Tables(6).Cell(1, 2).Range.Text, True)
		' 通信地址
		m_expert_address = compactString(doc.Tables(7).Cell(1, 2).Range.Text, True)
		' 联系电话
		m_expert_tel = compactString(doc.Tables(7).Cell(1, 4).Range.Text, True)
		' 对论文涉及内容的熟悉程度
		Dim arrMasterLevelText: arrMasterLevelText = Array("", "很熟悉(√)", "熟悉(√)", "一般(√)")
		For i = 1 To UBound(arrMasterLevelText)
			If doc.Tables(7).Cell(2, 1).Range.Find.Execute(arrMasterLevelText(i)) Then
				m_expert_master_level = i
				Exit For
			End If
		Next
		' 对学位论文的总体评价
		For i = 1 To 4
			If doc.Tables(11).Cell(2, i + 1).Range.Find.Execute("√") Then
				m_reviewlevel = i
				Exit For
			End If
		Next
		' 是否同意申请论文答辩
		Set rg = doc.Tables(11).Cell(3, 2).Range
		Dim j: j = 1
		For i = 1 To rg.Paragraphs.Count
			tmp = compactString(rg.Paragraphs(i).Range.Text, True)
			If Len(tmp) > 1 Then
				If rg.Paragraphs(i).Range.Find.Execute("√") Then
					m_reviewresult = j
					Exit For
				End If
				j = j + 1
			End If
		Next
		' 学位论文内容与申请学位专业的相关性
		m_correlation_level = 1
		extractMEMInfo = 1
	End Function
	' 提取 MPAcc 类型评阅书信息
	Private Function extractMPAccInfo(doc)
		Dim tmp
		' 送审时间
		m_review_time = DateValue(compactString(doc.Tables(2).Cell(9, 2).Range.Text, True))
		' 送审方式
		m_review_pattern = compactString(doc.Tables(2).Cell(8, 2).Range.Text, True)
		' 学号
		m_stu_no = compactString(doc.Tables(9).Cell(1, 4).Range.Text, True)
		' 评语
		Dim rg: Set rg = doc.Tables(9).Cell(2, 1).Range
		Dim i
		For i = 2 To rg.Sentences.Count
			tmp = tmp & rg.Sentences(i).Text
		Next
		m_comment = compactString(tmp, False)
		' 评阅专家信息
		' 专业技术职务
		m_expert_pro_duty = compactString(doc.Tables(4).Cell(1, 4).Range.Text, True)
		' 学科专长
		m_expert_expertise = compactString(doc.Tables(5).Cell(1, 2).Range.Text, True)
		' 单位名称（含院系）
		m_expert_workplace = compactString(doc.Tables(6).Cell(1, 2).Range.Text, True)
		' 通信地址
		m_expert_address = compactString(doc.Tables(7).Cell(1, 2).Range.Text, True)
		' 联系电话
		m_expert_tel = compactString(doc.Tables(7).Cell(1, 4).Range.Text, True)
		' 对论文涉及内容的熟悉程度
		Dim arrMasterLevelText: arrMasterLevelText = Array("", "很熟悉(√)", "熟悉(√)", "一般(√)")
		For i = 1 To UBound(arrMasterLevelText)
			If doc.Tables(7).Cell(2, 1).Range.Find.Execute(arrMasterLevelText(i)) Then
				m_expert_master_level = i
				Exit For
			End If
		Next
		' 对学位论文的总体评价
		For i = 1 To 4
			If doc.Tables(10).Cell(2, i + 1).Range.Find.Execute("√") Then
				m_reviewlevel = i
				Exit For
			End If
		Next
		' 是否同意申请论文答辩
		Set rg = doc.Tables(10).Cell(3, 2).Range
		Dim j: j = 1
		For i = 1 To rg.Paragraphs.Count
			tmp = compactString(rg.Paragraphs(i).Range.Text, True)
			If Len(tmp) > 1 Then
				If rg.Paragraphs(i).Range.Find.Execute("√") Then
					m_reviewresult = j
					Exit For
				End If
				j = j + 1
			End If
		Next
		' 学位论文内容与申请学位专业的相关性
		m_correlation_level = 1
		extractMPAccInfo = 1
	End Function
	' 提取 EMBA 类型评阅书信息
	' 与 MPAcc 类型相同
	Private Function extractEMBAInfo(doc)
		extractEMBAInfo=extractMPAccInfo(doc)
	End Function
	Private Function compactString(ByVal str, singleLine)
		str = Replace(Replace(str, Chr(7), ""), ChrW(160), "")
		If singleLine Then
			str = Replace(Replace(Replace(str, vbCr, ""), vbLf, ""), Chr(11), "")
		Else
			str = Replace(str, Chr(11), vbNewLine)
		End If
		compactString = str
	End Function
	Private Function getScore(ByVal str)
		getScore = Trim(Split(str, "分")(0))
	End Function

End Class
%>
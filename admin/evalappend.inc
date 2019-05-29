<%
' 审批材料导师意见追加工具类
Class EvalAppend
	Dim m_tutorname,m_date,m_eval_text
	Dim m_err_desc
	Public Property Let TutorName(val)
		m_tutorname=val
	End Property
	Public Property Let Date(val)
		m_date=val
	End Property
	Public Property Let EvalText(val)
		m_eval_text=val
	End Property
	Public Property Get ErrorDesc()
		ErrorDesc=m_err_desc
	End Property
	Public Function appendEval(filepath,tablefile,filetype)
		' 追加导师意见
		Dim wd,doc,rg,rg2,content
		Dim arrText,textid1,textid2
		arrText=Array("校内指导教师意见：（包括对申请人的学习情况、思想表现及论文的学术评语，科研工作能力和完成科研工作情况，以及是否同意申请学位论文答辩的意见）",_
									"校内导师签名：　",_
									"指导教师意见：（包括对申请人的业务学习、思想表现及学位论文的学术评语，科研工作能力和完成科研工作情况，以及是否同意申请学位论文答辩的意见）",_
									"导师签名：　")
		'On Error Resume Next
		Select Case filetype
		Case 1
			textid1=0
			textid2=1
		Case 2
			textid1=2
			textid2=3
		End Select
		Set wd=Server.CreateObject("Word.Application")
		Set doc=wd.Documents.Add(tablefile)
		Set rg=doc.Tables(3).Cell(12,1).Range
		content=arrText(textid1)&vbNewLine&vbNewLine&m_eval_text&vbNewLine&vbNewLine
		rg.Delete()
		rg.ParagraphFormat.Reset()
		rg.ParagraphFormat.LineUnitBefore=0.5
		rg.ParagraphFormat.LineUnitAfter=0.5
		rg.Paragraphs.Add()
		Set rg=rg.Paragraphs(1).Range
		rg.Font.Size=12
		rg.Text=content
		rg.InsertParagraphAfter()
		Set rg2=rg.Paragraphs(rg.Paragraphs.Count).Range
		rg2.Font.Size=12
		rg2.Text=arrText(textid2)&m_tutorname&vbNewLine&vbNewLine&Year(m_date)&"  年  "&Month(m_date)&"  月  "&Day(m_date)&"  日"
		rg2.ParagraphFormat.IndentFirstLineCharWidth 13
		doc.SaveAs filepath,16
		doc.Close(False)
		Set rg=Nothing
		Set rg2=Nothing
		Set doc=Nothing
		wd.Quit()
		Set wd=Nothing
		m_err_desc=Err.Description
		If Err.Number=0 Then
			appendEval=1
		Else
			appendEval=0
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
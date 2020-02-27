<%Response.Expires=-1%>
<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="appgen.inc"-->
<!--#include file="evalappend.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
Dim Upload:Set Upload=New ExtendedRequest
step=Request.QueryString("step")
thesisID=Request.QueryString("tid")
new_activity_id=Upload.Form("new_activity_id")
new_subject_ch=Upload.Form("new_subject_ch")
new_subject_en=Upload.Form("new_subject_en")
new_researchway_name=Upload.Form("new_researchway_name")
new_keywords_ch=Upload.Form("new_keywords_ch")
new_keywords_en=Upload.Form("new_keywords_en")
new_review_type=Upload.Form("new_review_type")
new_submit_review_time=Upload.Form("new_submit_review_time")
new_task_progress=Upload.Form("new_task_progress")
new_review_status=Upload.Form("new_review_status")
new_reproduct_ratio=Upload.Form("reproduct_ratio")
new_instruct_review_reproduct_ratio=Upload.Form("instruct_review_reproduct_ratio")
new_defence_result=Upload.Form("new_defence_result")
new_grant_degree_result=Upload.Form("new_grant_degree_result")
opr=Int(Upload.Form("opr"))
submittype=Upload.Form("submittype")
is_passed=submittype="pass"
eval_text=Upload.Form("eval_text")
Set detect_report=Upload.File("detect_report")
Set instruct_review_detect_report=Upload.File("instruct_review_detect_report")
activity_id=Upload.Form("In_ActivityId2")
teachtype_id=Upload.Form("In_TEACHTYPE_ID2")
class_id=Upload.Form("In_CLASS_ID2")
enter_year=Upload.Form("In_ENTER_YEAR2")
query_task_progress=Upload.Form("In_TASK_PROGRESS2")
query_review_status=Upload.Form("In_REVIEW_STATUS2")
finalFilter=Upload.Form("finalFilter2")
pageSize=Upload.Form("pageSize2")
pageNo=Upload.Form("pageNo2")
If Len(thesisID)=0 Or Not IsNumeric(thesisID) Or Not IsNumeric(opr) Then
	bError=True
	errdesc="参数无效。"
ElseIf submittype<>vbNullString And Not isMatched("[0-8]",opr,True) Then
	bError=True
	errdesc="操作无效。"
ElseIf submittype="unpass" And opr<=3 Or submittype<>vbNullString And (opr=4 Or opr=5 Or opr=6) Then
	If Len(eval_text)=0 Then
		bError=True
		errdesc="请填写意见（200-2000字）！"
	ElseIf Len(eval_text)>2000 Then
		bError=True
		errdesc="意见字数超出限制（2000字）！"
	End If
ElseIf Not (new_submit_review_time = vbNullString Or IsDate(new_submit_review_time)) Then
	bError=True
	errdesc="送审意见提交时间格式无效，正确格式为：年/月/日 时:分:秒！"
ElseIf new_reproduct_ratio<>vbNullString And Not IsNumeric(new_reproduct_ratio) Then
	bError=True
	errdesc="送检论文复制比输入无效，请输入 0-100 间的数字！"
ElseIf new_instruct_review_reproduct_ratio<>vbNullString And Not IsNumeric(new_instruct_review_reproduct_ratio) Then
	bError=True
	errdesc="教指委盲评论文复制比输入无效，请输入 0-100 间的数字！"
ElseIf Not isMatched("[0-4]",new_defence_result,True) Then
	bError=True
	errdesc="答辩成绩输入无效！"
ElseIf Not isMatched("[0-3]",new_grant_degree_result,True) Then
	bError=True
	errdesc="答辩表决结果设置无效！"
ElseIf detect_report.FileName<>vbNullString And new_reproduct_ratio=vbNullString Then
	bError=True
	errdesc="请填写送检论文复制比！"
ElseIf instruct_review_detect_report.FileName<>vbNullString And new_instruct_review_reproduct_ratio=vbNullString Then
	bError=True
	errdesc="请填写教指委盲评论文复制比！"
End If
If bError Then
%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	CloseRs rs
	CloseConn conn
	Response.End()
End If

If Len(new_reproduct_ratio) Then
	new_reproduct_ratio=Round(new_reproduct_ratio,4)
Else
	new_reproduct_ratio=Null
End If

Dim conn,rs,sql,sqlDetect,count
Connect conn
sql="SELECT * FROM Dissertations WHERE ID="&thesisID
GetRecordSet conn,rs,sql,count
If rs.EOF Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">数据库没有该论文记录！</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
  CloseRs rs
  CloseConn conn
	Response.End()
End If

If submittype=vbNullString Then
	opr=0
End If
review_status=rs("REVIEW_STATUS")
will_add_audit=False
Select Case opr
Case 1	'	 审核开题报告表
	filetypename="开题报告表/开题论文"
	If is_passed Then
		rs("TASK_PROGRESS")=tpTbl1Passed
	Else
		rs("TASK_PROGRESS")=tpTbl1Unpassed
	End If
	will_add_audit=True
	audit_type=auditTypeKtbg
Case 2	'  审核中期检查表
	filetypename="中期检查表/中期论文"
	If is_passed Then
		rs("TASK_PROGRESS")=tpTbl2Passed
	Else
		rs("TASK_PROGRESS")=tpTbl2Unpassed
	End If
	will_add_audit=True
	audit_type=auditTypeZqjcb
Case 3	'  审核预答辩申请
	filetypename="预答辩申请表/预答辩论文"
	If is_passed Then
		' 更新记录
		rs("TASK_PROGRESS")=tpTbl3Passed
	Else
		rs("TASK_PROGRESS")=tpTbl3Unpassed
	End If
	will_add_audit=True
	audit_type=auditTypeYdbyjs
Case 4	'  审核答辩材料
	filetypename="答辩审批材料"
	If is_passed Then
		rs("TASK_PROGRESS")=tpTbl4Passed
	Else
		rs("TASK_PROGRESS")=tpTbl4Unpassed
	End If
	will_add_audit=True
	audit_type=auditTypeSpcl
Case 5	'  同意/不同意送检送审操作
	filetypename="送检论文"
	author=Upload.Form("author")
	stuno=Upload.Form("stuno")
	tutorinfo=Upload.Form("tutorinfo")
	speciality=Upload.Form("speciality")
	faculty=Upload.Form("faculty")
	subject=Upload.Form("subject")
	If Not is_passed And (Len(author)=0 Or Len(stuno)=0 Or Len(tutorinfo)=0 Or Len(speciality)=0 Or Len(faculty)=0 _
	Or Len(subject)=0) Then
		bError=True
		errdesc="缺少必要的字段信息！"
	ElseIf review_status>=rsRefusedDetect Then
		bError=True
		errdesc="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		CloseRs rs
		CloseConn conn
		Response.End()
	End If
	' 更新记录
	If is_passed Then
		sql="SELECT dbo.getDetectResultCount("&thesisID&")"
		GetRecordSet conn,rsDetect,sql,count
		detect_count=rsDetect(0)
		CloseRs rsDetect
		rs("REVIEW_APP_EVAL")=eval_text
		rs("SUBMIT_REVIEW_TIME")=Now
		rs("REVIEW_STATUS")=rsAgreedDetect
		If detect_count>=1 Then
			eval_text="该生已对论文进行修改，并已经导师检查，同意二次检测。"
		Else
			eval_text="论文已检查，同意检测。"
		End If
	Else
		rs("REVIEW_STATUS")=rsRefusedDetect
	End If
	will_add_audit=True
	audit_type=auditTypeDetectReview
Case 6	'  同意/不同意送审操作
	filetypename="送审论文"
	author=Upload.Form("author")
	stuno=Upload.Form("stuno")
	tutorinfo=Upload.Form("tutorinfo")
	speciality=Upload.Form("speciality")
	faculty=Upload.Form("faculty")
	subject=Upload.Form("new_subject")
	reproduct_ratio=Upload.Form("reproduct_ratio")
	If Len(author)=0 Or Len(stuno)=0 Or Len(tutorinfo)=0 Or Len(speciality)=0 Or Len(faculty)=0 _
	Or Len(subject)=0 Or is_passed And Len(reproduct_ratio)=0 Then
		bError=True
		errdesc="缺少必要的字段信息！"
	ElseIf review_status>=rsRefusedReview Then
		bError=True
		errdesc="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		CloseRs rs
		CloseConn conn
		Response.End()
	End If
	If is_passed Then
		' 生成送审申请表
		Dim rag,audit_time
		audit_time=Now
		Randomize()
		filename=FormatDateTime(audit_time,1)&Int(Timer)&Int(Rnd()*999)&".docx"
		filepath=Server.MapPath("/PaperReview/tutor/export")&"\"&filename
		Set rag=New ReviewAppGen
		rag.Author=author
		rag.StuNo=stuno
		rag.TutorInfo=tutorinfo
		rag.Spec=speciality
		rag.Date=FormatDateTime(audit_time,1)
		rag.Subject=subject
		rag.EvalText=eval_text
		rag.ReproductRatio=reproduct_ratio
		bError=rag.generateApp(filepath)=0
		Set rag=Nothing
		rs("REVIEW_APP")=filename
		rs("REVIEW_STATUS")=rsAgreedReview
	Else
		rs("REVIEW_STATUS")=rsRefusedReview
	End If
	will_add_audit=True
	audit_type=auditTypeReviewApp
Case 7	' 评阅书审阅确认操作
	If review_status>=rsReviewEval Then
		bError=True
		errdesc="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		CloseRs rs
		CloseConn conn
		Response.End()
	End If
	' 更新记录
	rs("REVIEW_STATUS")=rsReviewEval
Case 8	'  提交答辩论文审核意见操作
	filetypename="答辩论文"
	If review_status>=rsRefusedDefence Then
		bError=True
		errdesc="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		CloseRs rs
		CloseConn conn
		Response.End()
	End If
	' 更新记录
	If is_passed Then
		eval_text="已审阅，同意答辩"
		rs("REVIEW_STATUS")=rsAgreedDefence
	Else
		eval_text="不同意答辩，请继续修改论文"
		rs("REVIEW_STATUS")=rsRefusedDefence
	End If
	will_add_audit=True
	audit_type=auditTypeDefence
End Select
If submittype=vbNullString Then
	' 更新表单信息
	Dim arrDetectFileFieldNames:arrDetectFileFieldNames=Array("THESIS_FILE","THESIS_FILE4")
	Dim arrDetectResultFieldNames:arrDetectResultFieldNames=Array("REPRODUCTION_RATIO","INSTRUCT_REVIEW_REPRODUCTION_RATIO")
	Dim arrReportFiles:arrReportFiles=Array(detect_report,instruct_review_detect_report)
	Dim arrNewDetectResults:arrNewDetectResults=Array(new_reproduct_ratio,new_instruct_review_reproduct_ratio)
	Dim i
	For i=0 To 1
		reportDir=getDateTimeId(Now)
		uploadPath=Server.MapPath("upload/report/"&reportDir)
		detectThesis=rs(arrDetectFileFieldNames(i))
		If arrReportFiles(i).FileName<>vbNullString Then
			ensurePathExists uploadPath
			destFile=generateDateTimeFilename(LCase(arrReportFiles(i).FileExt))
			arrReportFiles(i).SaveAs uploadPath&"\"&destFile
			sqlDetect="EXEC spSetDetectResultReport "&thesisID&","&toSqlString(detectThesis)&","&toSqlString(reportDir&"/"&destFile)&";"
		End If
		If Not IsNull(detectThesis) Then
			ratio=rs(arrDetectResultFieldNames(i))
			new_ratio=arrNewDetectResults(i)
			If Not IsNull(ratio) And new_ratio=vbNullString Then
				sqlDetect=sqlDetect&"EXEC spDeleteDetectResult "&thesisID&","&toSqlString(detectThesis)&";"
			ElseIf new_ratio <> vbNullString Then
				sqlDetect=sqlDetect&"EXEC spSetDetectResultRatio "&thesisID&","&toSqlString(detectThesis)&","&toSqlNumber(new_ratio)&";"
			End If
		End If
	Next
	
	If Len(new_activity_id)=0 Then
		rs("ActivityId")=Null
	Else
		rs("ActivityId")=new_activity_id
	End If
	If Len(new_defence_result)<>0 Then
		sql="UPDATE DefenceInfo SET DEFENCE_RESULT="&new_defence_result&" WHERE THESIS_ID="&thesisID
		conn.Execute sql
	End If
	If Len(new_grant_degree_result)<>0 Then
		rs("GRANT_DEGREE_RESULT")=new_grant_degree_result
	End If
	If Len(new_subject_ch)=0 Then
		rs("THESIS_SUBJECT")=Null
	Else
		rs("THESIS_SUBJECT")=new_subject_ch
	End If
	If Len(new_subject_en)=0 Then
		rs("THESIS_SUBJECT_EN")=Null
	Else
		rs("THESIS_SUBJECT_EN")=new_subject_en
	End If
	If Len(new_researchway_name)=0 Then
		rs("RESEARCHWAY_NAME")=Null
	Else
		rs("RESEARCHWAY_NAME")=new_researchway_name
	End If
	If Len(new_keywords_ch)=0 Then
		rs("KEYWORDS")=Null
	Else
		rs("KEYWORDS")=new_keywords_ch
	End If
	If Len(new_keywords_en)=0 Then
		rs("KEYWORDS_EN")=Null
	Else
		rs("KEYWORDS_EN")=new_keywords_en
	End If
	If Len(new_review_type)=0 Then
		rs("REVIEW_TYPE")=Null
	Else
		rs("REVIEW_TYPE")=new_review_type
	End If
	If Len(new_submit_review_time)=0 Then
		rs("SUBMIT_REVIEW_TIME")=Null
	Else
		rs("SUBMIT_REVIEW_TIME")=new_submit_review_time
	End If
	If Len(new_task_progress) Then
		rs("TASK_PROGRESS")=new_task_progress
	End If
	If Len(new_review_status) Then
		rs("REVIEW_STATUS")=new_review_status
	End If
End If
rs.Update()
If Len(sqlDetect) Then
	conn.Execute sqlDetect
End If
CloseRs rs
CloseConn conn

If will_add_audit Then
	' 插入审核记录
	addAuditRecord dissertation_id, filename, audit_type, audit_time, Session("TId"), is_passed, eval_text
End If
If opr=7 Then
	' 向学生发送修改论文通知邮件
	sendEmailToStudent thesisID,"",True,""
ElseIf opr<>0 Then
	' 向学生发送审核结果通知邮件
	sendEmailToStudent thesisID,filetypename,is_passed,eval_text
End If
%><form id="ret" action="thesisDetail.asp?tid=<%=thesisID%>" method="post">
<input type="hidden" name="In_ActivityId2" value="<%=activity_id%>">
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID2" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form>
<script type="text/javascript">
	alert("操作完成。");
	document.all.ret.submit();
</script>
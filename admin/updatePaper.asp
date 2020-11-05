<%Response.Expires=-1%>
<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/automation/ReviewApplicationFormWriter.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
Dim Upload:Set Upload=New ExtendedRequest
step=Request.QueryString("step")
paper_id=Request.QueryString("tid")
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
is_pass=submittype="agree"
comment=Upload.Form("comment")
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
If Len(paper_id)=0 Or Not IsNumeric(paper_id) Or Not IsNumeric(opr) Then
	bError=True
	errMsg="参数无效。"
ElseIf submittype<>vbNullString And Not isMatched("[0-9]",opr,True) Then
	bError=True
	errMsg="操作无效。"
ElseIf Not is_pass And Len(submittype) And isMatched("[12345689]",opr,True) Then
	If Len(comment)=0 Then
		bError=True
		errMsg="请填写意见（200-2000字）！"
	ElseIf Len(comment)>2000 Then
		bError=True
		errMsg="意见字数超出限制（2000字）！"
	End If
ElseIf Not (new_submit_review_time = vbNullString Or IsDate(new_submit_review_time)) Then
	bError=True
	errMsg="送审意见提交时间格式无效，正确格式为：年/月/日 时:分:秒！"
ElseIf new_reproduct_ratio<>vbNullString And Not IsNumeric(new_reproduct_ratio) Then
	bError=True
	errMsg="送检论文复制比输入无效，请输入 0-100 间的数字！"
ElseIf new_instruct_review_reproduct_ratio<>vbNullString And Not IsNumeric(new_instruct_review_reproduct_ratio) Then
	bError=True
	errMsg="教指委盲评论文复制比输入无效，请输入 0-100 间的数字！"
ElseIf Not isMatched("[0-4]",new_defence_result,True) Then
	bError=True
	errMsg="答辩成绩输入无效！"
ElseIf Not isMatched("[0-3]",new_grant_degree_result,True) Then
	bError=True
	errMsg="答辩表决结果设置无效！"
ElseIf detect_report.FileName<>vbNullString And new_reproduct_ratio=vbNullString Then
	bError=True
	errMsg="请填写送检论文复制比！"
ElseIf instruct_review_detect_report.FileName<>vbNullString And new_instruct_review_reproduct_ratio=vbNullString Then
	bError=True
	errMsg="请填写教指委盲评论文复制比！"
End If
If bError Then
%><body><center><font color=red size="4"><%=errMsg%></font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
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
ConnectDb conn
sql=Format("SELECT * FROM Dissertations WHERE ID={0}",paper_id)
GetRecordSet conn,rs,sql,count
If rs.EOF Then
%><body><center><font color=red size="4">数据库没有该论文记录！</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	CloseRs rs
	CloseConn conn
	Response.End()
End If

If submittype=vbNullString Then
	opr=0
End If
sql="SELECT TUTOR_ID FROM ViewDissertations WHERE ID="&paper_id
Set rsTutor=ExecQuery(conn, sql)("rs")
tutor_id=rsTutor("TUTOR_ID")
CloseRs rsTutor

audit_time=Now
review_status=rs("REVIEW_STATUS")
will_add_audit=False
Select Case opr
Case 1	'	 审核开题报告表
	file_type_name="开题报告表/开题论文"
	audit_file=rs("TABLE_FILE1")
	If is_pass Then
		rs("TASK_PROGRESS")=tpTbl1Passed
	Else
		rs("TASK_PROGRESS")=tpTbl1Unpassed
	End If
	will_add_audit=True
	audit_type=auditTypeKtbgb
Case 2	'  审核中期考核表
	file_type_name="中期考核表/中期论文"
	audit_file=rs("TABLE_FILE2")
	If is_pass Then
		rs("TASK_PROGRESS")=tpTbl2Passed
	Else
		rs("TASK_PROGRESS")=tpTbl2Unpassed
	End If
	will_add_audit=True
	audit_type=auditTypeZqkhb
Case 3	'  审核预答辩意见书
	file_type_name="预答辩意见书/预答辩论文"
	audit_file=rs("TABLE_FILE3")
	If is_pass Then
		rs("TASK_PROGRESS")=tpTbl3Passed
	Else
		rs("TASK_PROGRESS")=tpTbl3Unpassed
	End If
	will_add_audit=True
	audit_type=auditTypeYdbyjs
Case 4	'  审核答辩材料
	file_type_name="答辩审批材料"
	audit_file=rs("TABLE_FILE4")
	If is_pass Then
		rs("TASK_PROGRESS")=tpTbl4Passed
	Else
		rs("TASK_PROGRESS")=tpTbl4Unpassed
	End If
	will_add_audit=True
	audit_type=auditTypeSpclb
Case 5	'  同意/不同意送检送审操作
	file_type_name="送检论文"
	audit_file=rs("THESIS_FILE")
	author=Upload.Form("author")
	stuno=Upload.Form("stuno")
	tutorinfo=Upload.Form("tutorinfo")
	faculty=Upload.Form("faculty")
	subject=Upload.Form("subject")
	If Not is_pass And (Len(author)=0 Or Len(stuno)=0 Or Len(tutorinfo)=0 Or Len(faculty)=0 _
	Or Len(subject)=0) Then
		bError=True
		errMsg="缺少必要的字段信息！"
	ElseIf review_status>=rsRefusedDetect Then
		bError=True
		errMsg="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		CloseRs rs
		CloseConn conn
		showErrorPage errMsg, "提示"
	End If
	If is_pass Then
		sql="SELECT dbo.getDetectResultCount("&paper_id&")"
		GetRecordSet conn,rsDetect,sql,count
		detect_count=rsDetect(0)
		CloseRs rsDetect
		rs("REVIEW_APP_EVAL")=comment
		rs("SUBMIT_REVIEW_TIME")=Now
		rs("REVIEW_STATUS")=rsAgreedDetect
		If detect_count>=1 Then
			comment="该生已对论文进行修改，并已经导师检查，同意二次检测。"
		Else
			comment="论文已检查，同意检测。"
		End If
	Else
		rs("REVIEW_STATUS")=rsRefusedDetect
	End If
	will_add_audit=True
	audit_type=auditTypeDetectReview
Case 7	' 评阅书审阅确认操作
	audit_file=rs("REVIEW_FILE")
	If review_status>=rsReviewEval Then
		bError=True
		errMsg="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		CloseRs rs
		CloseConn conn
		showErrorPage errMsg, "提示"
	End If
	rs("REVIEW_STATUS")=rsReviewEval
Case 8	'  提交答辩论文审核意见操作
	file_type_name="答辩论文"
	audit_file=rs("THESIS_FILE3")
	If review_status>=rsRefusedDefence Then
		bError=True
		errMsg="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		CloseRs rs
		CloseConn conn
		showErrorPage errMsg, "提示"
	End If
	' 更新记录
	If is_pass Then
		comment="已审阅，同意答辩"
		rs("REVIEW_STATUS")=rsAgreedDefence
	Else
		comment="不同意答辩，请继续修改论文"
		rs("REVIEW_STATUS")=rsRefusedDefence
	End If
	will_add_audit=True
	audit_type=auditTypeDefence
Case 9	'  提交教指委盲评论文审核意见操作
	file_type_name="教指委盲评论文"
	audit_file=rs("THESIS_FILE4")
	If review_status>=rsRefusedInstructReview Then
		bError=True
		errMsg="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		CloseRs rs
		CloseConn conn
		showErrorPage errMsg, "提示"
	End If
	' 更新记录
	If is_pass Then
		'comment="已审阅，同意答辩"
		rs("REVIEW_STATUS")=rsAgreedInstructReview
	Else
		'comment="不同意答辩，请继续修改论文"
		rs("REVIEW_STATUS")=rsRefusedInstructReview
	End If
	will_add_audit=True
	audit_type=auditTypeInstructReviewDetect
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
		uploadPath=Server.MapPath(resolvePath(uploadBasePath(usertypeAdmin,"detect_report"),reportDir))
		detectThesis=rs(arrDetectFileFieldNames(i))
		If arrReportFiles(i).FileName<>vbNullString Then
			ensurePathExists uploadPath
			destFile=generateDateTimeFilename(LCase(arrReportFiles(i).FileExt))
			arrReportFiles(i).SaveAs resolvePath(uploadPath,destFile)
			sqlDetect="EXEC spSetDetectResultReport "&paper_id&","&toSqlString(detectThesis)&","&toSqlString(resolvePath(reportDir,destFile))&";"
		End If
		If Not IsNull(detectThesis) Then
			ratio=rs(arrDetectResultFieldNames(i))
			new_ratio=arrNewDetectResults(i)
			If new_ratio=vbNullString Then new_ratio=ratio
			If IsNull(new_ratio) Then new_ratio=0
			If IsNull(ratio) Then
				sqlDetect=sqlDetect&"EXEC spSetDetectResultRatio "&paper_id&","&toSqlString(detectThesis)&","&toSqlNumber(new_ratio)&";"
			ElseIf CDbl(ratio)<>CDbl(new_ratio) Then
				sqlDetect=sqlDetect&"EXEC spSetDetectResultRatio "&paper_id&","&toSqlString(detectThesis)&","&toSqlNumber(new_ratio)&";"
			End If
		End If
	Next
	
	If Len(new_activity_id)=0 Then
		rs("ActivityId")=Null
	Else
		rs("ActivityId")=new_activity_id
	End If
	If Len(new_defence_result)<>0 Then
		sql="UPDATE DefenceInfo SET DEFENCE_RESULT="&new_defence_result&" WHERE THESIS_ID="&paper_id
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
	If Len(new_task_progress) And rs("TASK_PROGRESS")<>Int(new_task_progress) Then
		rs("TASK_PROGRESS")=new_task_progress
		update_task_progress=True
	End If
	If Len(new_review_status) And rs("REVIEW_STATUS")<>Int(new_review_status) Then
		rs("REVIEW_STATUS")=new_review_status
		update_review_status=True
	End If
End If
rs.Update()
If Len(sqlDetect) Then
	conn.Execute sqlDetect
End If
CloseRs rs
If update_task_progress Or update_review_status Then
	sql=Format("SELECT STU_NAME,TASK_PROGRESS_NAME,REVIEW_STATUS_NAME FROM ViewDissertations WHERE ID={0}",paper_id)
	Set rs = conn.Execute(sql)
	If update_task_progress Then
		writeLog Format("教务员[{0}]修改学生[{1}]的论文表格审核进度为[{2}]。",Session("name"),rs(0),rs(1))
	End If
	If update_review_status Then
		writeLog Format("教务员[{0}]修改学生[{1}]的论文评阅进度为[{2}]。",Session("name"),rs(0),rs(2))
	End If
	CloseRs rs
End If
CloseConn conn

If will_add_audit Then
	' 插入审核记录
	addAuditRecord paper_id, audit_file, audit_type, audit_time, tutor_id, is_pass, comment
End If
If opr=7 Then
	' 向学生发送修改论文通知邮件
	sendEmailToStudent paper_id, "", True, ""
ElseIf opr<>0 Then
	' 向学生发送审核结果通知邮件
	sendEmailToStudent paper_id, file_type_name, is_pass, comment
End If
%><form id="ret" action="paperDetail.asp?tid=<%=paper_id%>" method="post">
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
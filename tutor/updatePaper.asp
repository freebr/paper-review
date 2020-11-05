<%Response.Expires=-1%>
<!--#include file="../inc/global.inc"-->
<!--#include file="../inc/automation/ReviewApplicationFormWriter.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")
step=Request.QueryString("step")
paper_id=Request.QueryString("tid")
opr=Request.Form("opr")
submittype=Request.Form("submittype")
is_pass=submittype="agree"
comment=Request.Form("comment")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
spec_id=Request.Form("In_SPECIALITY_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
If Len(paper_id)=0 Or Not IsNumeric(paper_id) Then
	bError=True
	errMsg="参数无效。"
ElseIf Not isMatched("[1-9]",opr,True) Then
	bError=True
	errMsg="操作无效。"
ElseIf Not is_pass And isMatched("[12345689]",opr,True) Then
	If Len(comment)=0 Then
		bError=True
		errMsg="请填写意见（200-2000字）！"
	ElseIf Len(comment)>2000 Then
		bError=True
		errMsg="意见字数超出限制（2000字）！"
	End If
End If
If bError Then
	CloseRs rs
	CloseConn conn
	showErrorPage errMsg, "提示"
End If

Dim conn,sql,ret,rs,count
ConnectDb conn
sql="SELECT * FROM ViewDissertations_instruct WHERE ID=?"
Set ret=ExecQuery(conn,sql,CmdParam("ID",adInteger,4,paper_id))
Set rs=ret("rs")
If rs.EOF Then
  	CloseRs rs
  	CloseConn conn
	showErrorPage "数据库没有该论文记录！", "提示"
End If

Dim section_id,is_comment,instruct_member
teacher_id=Session("TId")
If opr=11 Then
	section_id=sectionInstructReview
	is_comment=Array(rs("IsComment1"),rs("IsComment2"))
	If rs("INSTRUCT_MEMBER1")=teacher_id Then
		instruct_member=0
	ElseIf rs("INSTRUCT_MEMBER2")=teacher_id Then
		instruct_member=1
	End If
Else
	section_id=sectionAudit
End If

Dim section_access_info
Set section_access_info=getSectionAccessibilityInfo(rs("ActivityId"),rs("TEACHTYPE_ID"),section_id)
If Not section_access_info("accessible") Then
	CloseRs rs
  	CloseConn conn
	showErrorPage section_access_info("tip"), "提示"
End If

CloseRs rs
sql="SELECT * FROM Dissertations WHERE ID="&paper_id
GetRecordSet conn,rs,sql,count
audit_time=Now
review_status=rs("REVIEW_STATUS")
will_add_audit=False
will_notify=True
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
	file_type_name="送检论文和送审论文"
	audit_file=rs("THESIS_FILE")
	author=Request.Form("author")
	stuno=Request.Form("stuno")
	tutorinfo=Request.Form("tutorinfo")
	speciality=Request.Form("speciality")
	faculty=Request.Form("faculty")
	subject=Request.Form("subject")
	If Not is_pass And (Len(author)=0 Or Len(stuno)=0 Or Len(tutorinfo)=0 Or Len(speciality)=0 Or Len(faculty)=0 _
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
		If detect_count>=1 Then	' 二次检测
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
	If is_pass Then
		'comment="已审阅，同意答辩"
		rs("REVIEW_STATUS")=rsAgreedDefence
	Else
		'comment="不同意答辩，请继续修改论文"
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
	If is_pass Then
		rs("REVIEW_STATUS")=rsAgreedInstructReview
	Else
		rs("REVIEW_STATUS")=rsRefusedInstructReview
	End If
	will_add_audit=True
	audit_type=auditTypeInstructReviewDetect
Case 11	'  提交教指委盲评论文修改意见操作
	file_type_name="教指委盲评论文"
	audit_file=rs("THESIS_FILE4")
	If review_status>=rsInstructEval Then
		bError=True
		errMsg="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		CloseRs rs
		CloseConn conn
		showErrorPage errMsg, "提示"
	End If
	
	If is_comment(1-instruct_member) Then
		rs("REVIEW_STATUS")=rsInstructEval
	Else
		will_notify=False
	End If
	will_add_audit=True
	audit_type=auditTypeInstructReview
End Select
rs.Update()
CloseRs rs
CloseConn conn

If will_add_audit Then
	' 插入审核记录
	addAuditRecord paper_id, audit_file, audit_type, audit_time, is_pass, comment
End If
If opr=7 Then
	' 向学生发送评阅意见确认通知邮件
	sendEmailToStudent paper_id, "", True, ""
ElseIf will_notify Then
	' 向学生发送审核结果通知邮件
	sendEmailToStudent paper_id, file_type_name, is_pass, comment
End If
updateActiveTime teacher_id
%><form id="ret" action="paperDetail.asp?tid=<%=paper_id%>" method="post">
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID2" value="<%=spec_id%>" />
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
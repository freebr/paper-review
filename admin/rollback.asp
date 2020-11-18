<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
paper_id=Request.Form("tid")
usertype=Request.Form("user")
opr=Request.Form("rollback_opr")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
class_id=Request.Form("In_CLASS_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
If Len(paper_id)=0 Or Not IsNumeric(paper_id) Or Len(usertype)=0 Or Not IsNumeric(usertype) Or Len(opr)=0 Or Not IsNumeric(opr) Then
%><body><center><font color=red size="4">参数无效。</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End()
End If

Dim conn,rs,sql,sqlDetect,count
ConnectDb conn
sql="SELECT * FROM Dissertations WHERE ID="&paper_id
GetRecordSet conn,rs,sql,count
If rs.EOF Then
%><body><center><font color=red size="4">数据库没有该论文记录！</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	CloseRs rs
	CloseConn conn
	Response.End()
End If

Dim detect_count
sql="SELECT DETECT_COUNT FROM ViewDissertations WHERE ID="&paper_id
GetRecordSet conn,rsDetect,sql,count
detect_count=rsDetect(0)
CloseRs rsDetect

Select Case usertype
Case 0	' 撤销学生上传操作
	Select Case opr
	Case 0
		rs("TABLE_FILE1")=Null
		'rs("TASK_EVAL")=Null
		If IsNull(rs("TBL_THESIS_FILE1")) Then
			rs("TASK_PROGRESS")=0
		Else
			rs("TASK_PROGRESS")=tpTbl1Uploaded
		End If
		rs("REVIEW_STATUS")=0
	Case 1
		rs("TBL_THESIS_FILE1")=Null
		'rs("TASK_EVAL")=Null
		If IsNull(rs("TABLE_FILE1")) Then
			rs("TASK_PROGRESS")=0
		Else
			rs("TASK_PROGRESS")=tpTbl1Uploaded
		End If
		rs("REVIEW_STATUS")=0
	Case 2
		rs("TABLE_FILE2")=Null
		'rs("TASK_EVAL")=Null
		If IsNull(rs("TBL_THESIS_FILE2")) Then
			rs("TASK_PROGRESS")=tpTbl1Passed
		Else
			rs("TASK_PROGRESS")=tpTbl2Uploaded
		End If
		rs("REVIEW_STATUS")=0
	Case 3
		rs("TBL_THESIS_FILE2")=Null
		'rs("TASK_EVAL")=Null
		If IsNull(rs("TABLE_FILE2")) Then
			rs("TASK_PROGRESS")=tpTbl1Passed
		Else
			rs("TASK_PROGRESS")=tpTbl2Uploaded
		End If
		rs("REVIEW_STATUS")=0
	Case 4
		rs("TABLE_FILE3")=Null
		'rs("TASK_EVAL")=Null
		If IsNull(rs("TBL_THESIS_FILE3")) Then
			rs("TASK_PROGRESS")=tpTbl2Passed
		Else
			rs("TASK_PROGRESS")=tpTbl3Uploaded
		End If
		rs("REVIEW_STATUS")=0
	Case 5
		rs("TBL_THESIS_FILE3")=Null
		'rs("TASK_EVAL")=Null
		If IsNull(rs("TABLE_FILE3")) Then
			rs("TASK_PROGRESS")=tpTbl2Passed
		Else
			rs("TASK_PROGRESS")=tpTbl3Uploaded
		End If
		rs("REVIEW_STATUS")=0
	Case 6
		rs("THESIS_FILE")=Null
		'rs("DETECT_APP_EVAL")=Null
		rs("TASK_PROGRESS")=tpTbl3Passed
		rs("REVIEW_STATUS")=0
	Case 7
		rs("THESIS_FILE2")=Null
		'rs("REVIEW_APP_EVAL")=Null
		rs("REVIEW_STATUS")=rsAgreedDetect
	Case 8
		rs("THESIS_FILE3")=Null
		'rs("TUTOR_MODIFY_EVAL")=Null
		rs("REVIEW_STATUS")=rsReviewEval
	Case 9
		rs("THESIS_FILE4")=Null
		rs("REVIEW_STATUS")=rsDefenceEval
	Case 10
		rs("THESIS_FILE5")=Null
		rs("REVIEW_STATUS")=rsInstructEval
	Case 11
		rs("TABLE_FILE4")=Null
		'rs("TASK_EVAL")=Null
		rs("TASK_PROGRESS")=tpTbl3Passed
	End Select
Case 1	' 撤销专家评阅操作

	Dim review_result(2),review_time(1),review_level(1)
	If Not IsNull(rs("REVIEW_RESULT")) Then
		arr=Split(rs("REVIEW_RESULT"),",")
		For i=0 To UBound(arr)
			review_result(i)=Int(arr(i))
		Next
	End If
	If IsNull(rs("REVIEWER_EVAL_TIME")) Then
		For i=0 To 1
			review_level(i)=0
		Next
	Else
		arr4=Split(rs("REVIEWER_EVAL_TIME"),",")
		arr5=Split(rs("REVIEW_LEVEL"),",")
		For i=0 To 1
			review_time(i)=arr4(i)
			review_level(i)=Int(arr5(i))
		Next
	End If
	
	review_result(opr)="5"
	review_level(opr)="0"
	review_time(opr)=""
	finalresult="6"
	review_result(2)=finalresult
	
	' 更新记录
	rs("REVIEW_RESULT")=ArrayJoin(review_result,",")
	rs("REVIEW_LEVEL")=ArrayJoin(review_level,",")
	rs("REVIEWER_EVAL_TIME")=ArrayJoin(review_time,",")
	' TODO: 删除 ReviewRecords 表相应记录
	rs("REVIEW_STATUS")=rsMatchedReviewer
	
Case 2	' 撤销导师审核操作
	Select Case opr
	Case 0
		rs("TASK_EVAL")=Null
		Select Case rs("TASK_PROGRESS")
		Case tpTbl1Unpassed,tpTbl1Passed
			rs("TASK_PROGRESS")=tpTbl1Uploaded
			rs("REVIEW_STATUS")=0
		Case tpTbl2Unpassed,tpTbl2Passed
			rs("TASK_PROGRESS")=tpTbl2Uploaded
			rs("REVIEW_STATUS")=0
		Case tpTbl3Unpassed,tpTbl3Passed
			rs("TASK_PROGRESS")=tpTbl3Uploaded
			rs("REVIEW_STATUS")=0
		Case tpTbl4Unpassed,tpTbl4Passed
			rs("TASK_PROGRESS")=tpTbl4Uploaded
		End Select
	Case 1
		rs("DETECT_APP_EVAL")=Null
		rs("REVIEW_STATUS")=rsDetectPaperUploaded
	Case 2
		rs("REVIEW_APP")=Null
		rs("REVIEW_APP_EVAL")=Null
		rs("SUBMIT_REVIEW_TIME")=Null
		rs("DETECT_APP_EVAL")=Null
		rs("REVIEW_STATUS")=rsDetectPaperUploaded
	Case 3
		rs("TUTOR_MODIFY_EVAL")=Null
		rs("REVIEW_STATUS")=rsDefencePaperUploaded
	End Select
Case 3	' 撤销教务员操作
	Select Case opr
	Case 0	' 撤销送检操作
		sqlDetect="EXEC spDeleteDetectResult "&paper_id&";"
		If Not IsNull(rs("THESIS_FILE")) Then
			sqlDetect=sqlDetect&"EXEC spAddDetectResult "&paper_id&","&toSqlString(rs("THESIS_FILE"))&",NULL,NULL,NULL,1;"
		End If
		rs("REVIEW_STATUS")=rsAgreedDetect
	Case 1	' 撤销匹配评阅专家操作
		rs("REVIEW_STATUS")=rsAgreedReview
	Case 2	' 撤销导入答辩安排操作
		sql="DELETE FROM DefenceInfo WHERE THESIS_ID="&paper_id
		conn.Execute sql
	Case 3	' 撤销导入答辩委员会修改意见操作
		sql="UPDATE DefenceInfo SET DEFENCE_EVAL=NULL WHERE THESIS_ID="&paper_id
		conn.Execute sql
		rs("DEFENCE_MODIFY_EVAL")=Null	' 旧字段
		rs("REVIEW_STATUS")=rsAgreedDefence
	Case 4	' 撤销匹配教指委委员操作
		rs("REVIEW_STATUS")=rsAgreedInstructReview
	Case 5	' 撤销第一位教指委委员的修改意见
		audit_info=getAuditInfo(paper_id,rs("THESIS_FILE4"), auditTypeInstructReview)
		If Not IsEmpty(audit_info(0)("AuditorName")) Then
			addAuditRecord paper_id, rs("THESIS_FILE4"), auditTypeInstructReview, audit_info(0)("AuditTime"), rs("INSTRUCT_MEMBER1"), True, Null
		End If
		rs("REVIEW_STATUS")=rsMatchedInstructMember
	Case 6	' 撤销第二位教指委委员的修改意见
		audit_info=getAuditInfo(paper_id,rs("THESIS_FILE4"), auditTypeInstructReview)
		If UBound(audit_info)>=1 Then
			addAuditRecord paper_id, rs("THESIS_FILE4"), auditTypeInstructReview, audit_info(1)("AuditTime"), rs("INSTRUCT_MEMBER2"), True, Null
		End If
		rs("REVIEW_STATUS")=rsMatchedInstructMember
	Case 7	' 撤销导入学院学位评定分会修改意见操作
		rs("DEGREE_MODIFY_EVAL")=Null	' 旧字段
	End Select
End Select

If rs("REVIEW_STATUS")<rsMatchedReviewer Then
	rs("REVIEWER1")=Null
	rs("REVIEWER2")=Null
	rs("REVIEW_RESULT")="5,5,6"
	rs("REVIEW_LEVEL")="0,0"
	rs("REVIEWER_EVAL_TIME")=Null
	' TODO: 删除 ReviewRecords 表相应记录
End If
If rs("REVIEW_STATUS")<rsMatchedInstructMember Then
	rs("INSTRUCT_MEMBER1")=Null
	rs("INSTRUCT_MEMBER2")=Null
	' TODO: 删除 AuditRecords 表相应记录
End If
rs.Update()
If Len(sqlDetect) Then
	conn.Execute sqlDetect
End If
CloseRs rs
CloseConn conn
%><form id="ret" action="paperDetail.asp?tid=<%=paper_id%>" method="post">
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
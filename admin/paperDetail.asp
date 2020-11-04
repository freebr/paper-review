﻿<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
step=Request.QueryString("step")
paper_id=Request.QueryString("tid")
activity_id=Request.Form("In_ActivityId2")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
class_id=Request.Form("In_CLASS_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
If Len(paper_id)=0 Or Not IsNumeric(paper_id) Then
	showErrorPage "参数无效。", "提示"
End If

Connect conn
sql="SELECT * FROM ViewDissertations_admin WHERE ID="&paper_id
GetRecordSet conn,rs,sql,count
If count=0 Then
	CloseRs rs
	CloseConn conn
	showErrorPage "数据库没有该论文记录！", "提示"
End If

Dim review_status,numReviewed,review_result(2),review_file(1),review_time(1),review_level(1)
Dim rsDetect

sql="SELECT * FROM DetectResults WHERE DissertationId="&paper_id
GetRecordSet conn,rsDetect,sql,count

activity_id=rs("ActivityId")
stu_type=rs("TEACHTYPE_ID")
If stu_type=5 Then
	reviewfile_type=2
Else
	reviewfile_type=1
End If
tutor_id=rs("TUTOR_ID")
review_app=rs("REVIEW_APP")
review_type=rs("REVIEW_TYPE")
task_progress=rs("TASK_PROGRESS")
review_status=rs("REVIEW_STATUS")
stat_text1=rs("STAT_TEXT1")
stat_text2=rs("STAT_TEXT2")
detect_count=rs("DETECT_COUNT")
reproduct_ratio=rs("REPRODUCTION_RATIO")
instruct_review_reproduct_ratio=rs("INSTRUCT_REVIEW_REPRODUCTION_RATIO")
has_defence_plan=Not IsNull(rs("DEFENCE_MEMBER"))
defence_result=rs("DEFENCE_RESULT")
grant_degree_result=rs("GRANT_DEGREE_RESULT")
opr=0
Select Case task_progress
Case tpNone
Case tpTbl1Uploaded:opr=1
Case tpTbl2Uploaded:opr=2
Case tpTbl3Uploaded:opr=3
Case tpTbl4Uploaded:opr=4
End Select
Select Case review_status
Case rsDetectPaperUploaded:opr=5
Case rsMatchedReviewer:opr=10
Case rsReviewed:opr=7
Case rsDefencePaperUploaded:opr=8
Case rsInstructReviewPaperUploaded:opr=9
End Select
If review_status=0 Then
	stat=stat_text1
ElseIf task_progress>=tpTbl4Uploaded Then
	stat=stat_text1&"，"&stat_text2
Else
	stat=stat_text2
End If

Dim table_file(4)
For i=1 To 4
	table_file(i)=rs("TABLE_FILE"&i)
Next
If Not IsNull(rs("THESIS_FILE")) Then
	thesis_file=rs("THESIS_FILE")
End If
If Not IsNull(rs("THESIS_FILE2")) Then
	thesis_file_review=rs("THESIS_FILE2")
End If
If Not IsNull(rs("THESIS_FILE3")) Then
	thesis_file_modified=rs("THESIS_FILE3")
End If
If Not IsNull(rs("THESIS_FILE4")) Then
	thesis_file_instruct_review=rs("THESIS_FILE4")
End If
If Not IsNull(rs("THESIS_FILE5")) Then
	thesis_file_final=rs("THESIS_FILE5")
End If
If Not IsNull(rs("REVIEW_RESULT")) Then
	arr=Split(rs("REVIEW_RESULT"),",")
	For i=0 To UBound(arr)
		review_result(i)=Int(arr(i))
	Next
End If
If Not IsNull(rs("REVIEWER_EVAL_TIME")) Then
	arr=Array(rs("ReviewFile1"),rs("ReviewFile2"))
	arr2=Split(rs("REVIEWER_EVAL_TIME"),",")
	For i=0 To 1
		review_file(i)=arr(i)
		review_time(i)=arr2(i)
	Next
End If
Select Case step
Case vbNullString	' 论文详情页面
	Dim tutor_modify_eval_title
	arrActionUrlList=Array("?tid="&paper_id&"&step=2","updatePaper.asp?tid="&paper_id)
	Select Case opr
	Case 1,2,3
		actionUrl1=arrActionUrlList(0)
		actionUrl2=arrActionUrlList(1)
	Case 4,5,6,8,9
		actionUrl1=arrActionUrlList(0)
		actionUrl2=actionUrl1
	Case 7
		actionUrl1=arrActionUrlList(1)
		actionUrl2=actionUrl1
	Case 10
		actionUrl1="extra/paperDetail.asp?tid="&paper_id&"&rev=0&step=2"
		actionUrl2="extra/paperDetail.asp?tid="&paper_id&"&rev=1&step=2"
	End Select
	If review_status>=rsAgreedDefence Then
		tutor_modify_eval_title="导师同意答辩意见"
	ElseIf review_status=rsRefusedDefence Then
		tutor_modify_eval_title="导师不同意答辩意见"
	Else
		tutor_modify_eval_title="导师对答辩论文的意见"
	End If

	notice_text = getNoticeText(rs("TEACHTYPE_ID"), "review_result_desc")

	sql="SELECT * FROM ReviewTypes WHERE LEN(THESIS_FORM)>0 AND TEACHTYPE_ID="&stu_type
	GetRecordSetNoLock conn,rsRevType,sql,count
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>查看论文信息</title>
<% useStylesheet "admin", "jeasyui" %>
<% useScript "jquery", "jeasyui", "common", "*paper" %>
</head>
<body>
<h1>专业硕士论文详情<h2>论文当前状态：【<%=stat%>】</h2></h1>
<div class="box">
	<form id="fmDetail" action="updatePaper.asp?tid=<%=paper_id%>" enctype="multipart/form-data" method="post">
		<table class="form"><tr><td>
		<input type="button" id="btnupload" value="上传表格/论文文件..." onclick="submitForm(this.form,'uploadPaperFile.asp?tid=<%=paper_id%>')" />
		</td></tr></table>
		<fieldset>
			<legend>论文基本情况</legend>
			<table class="form">
			<tr><td class="field-name">评阅活动：</td><td><input class="easyui-combobox" id="activity_id" name="new_activity_id" /></td></tr>
			<tr><td class="field-name">论文题目：</td><td><input type="text" class="txt full-width" name="new_subject_ch" size="95%" value="<%=rs("THESIS_SUBJECT")%>" /></td></tr>
			<tr><td class="field-name">（英文）：</td><td><input type="text" class="txt full-width" name="new_subject_en" size="85%" value="<%=rs("THESIS_SUBJECT_EN")%>" /></td></tr>
			<tr><td class="field-name">作者姓名：</td><td><%=rs("STU_NAME")%><a class="open-window" href="#" onclick="return showStudentProfile(<%=rs("STU_ID")%>,0)">查看学生资料</a></td></tr>
			<tr><td class="field-name">学号：</td><td><input type="text" class="txt full-width" name="stuno" size="15" value="<%=rs("STU_NO")%>" readonly /></td></tr>
			<tr><td class="field-name">学位类别：</td><td><input type="text" class="txt full-width" name="degreename" size="10" value="<%=rs("TEACHTYPE_NAME")%>" readonly /></td></tr>
			<tr><td class="field-name">指导教师：</td><td><input type="text" class="txt full-width" name="tutorname" size="95%" value="<%=rs("TUTOR_NAME")%>" readonly /></td></tr><%
	If reviewfile_type=2 Then %>
			<tr><td class="field-name">领域名称：</td><td><input type="text" class="txt full-width" name="speciality" size="95%" value="<%=rs("SPECIALITY_NAME")%>" readonly /></td></tr><%
	End If %>
			<tr><td class="field-name">研究方向：</td><td><input type="text" class="txt full-width" name="new_researchway_name" size="95%" value="<%=rs("RESEARCHWAY_NAME")%>" /></td></tr>
			<tr><td class="field-name">论文关键词：</td><td><input type="text" class="txt full-width" name="new_keywords_ch" size="85%" value="<%=rs("KEYWORDS")%>" /></td></tr>
			<tr><td class="field-name">（英文）：</td><td><input type="text" class="txt full-width" name="new_keywords_en" size="85%" value="<%=rs("KEYWORDS_EN")%>" /></td></tr>
			<tr><td class="field-name">院系名称：</td><td><input type="text" class="txt full-width" name="faculty" size="30%" value="工商管理学院" readonly /></td></tr>
			<tr><td class="field-name">班级：</td><td><input type="text" class="txt full-width" name="class" size="51%" value="<%=rs("CLASS_NAME")%>" readonly /></td></tr><%
	If Not IsNull(rs("REVIEW_TYPE")) Then %>
			<tr><td class="field-name">论文形式：</td><td><select id="review_type" name="new_review_type" style="width:350px"><%
		Do While Not rsRevType.EOF
			%><option value="<%=rsRevType("ID")%>"<% If review_type=rsRevType("ID") Then %> selected<% End If %>><%=rsRevType("THESIS_FORM")%></option><%
			rsRevType.MoveNext()
		Loop
			%></select>
			</td></tr><%
	End If
	If review_status>=rsDetectPaperUploaded Then %>
			<tr><td class="field-name">送检论文文字复制比：</td><td><input type="text" class="txt" name="reproduct_ratio" value="<%=toNumericString(reproduct_ratio)%>" />%</td></tr><%
	End If
	If False And review_status>=rsInstructReviewPaperUploaded Then %>
			<tr><td class="field-name">教指委盲评论文文字复制比：</td><td><input type="text" class="txt" name="instruct_review_reproduct_ratio" value="<%=toNumericString(instruct_review_reproduct_ratio)%>" />%</td></tr><%
	End If
	If Len(table_file(1)) Then %>
			<tr><td class="field-name">开题报告表：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=1" target="_blank">点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,0,0)">撤销</a></td></tr><%
	End If
	If Not IsNull(rs("TBL_THESIS_FILE1")) Then %>
			<tr><td class="field-name">开题论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=2" target="_blank">点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,0,1)">撤销</a></td></tr><%
	End If
	If Len(table_file(2)) Then %>
			<tr><td class="field-name">中期考核表：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=3" target="_blank">点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,0,2)">撤销</a></td></tr><%
	End If
	If Not IsNull(rs("TBL_THESIS_FILE2")) Then %>
			<tr><td class="field-name">中期论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=4" target="_blank">点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,0,3)">撤销</a></td></tr><%
	End If
	If Len(table_file(3)) Then %>
			<tr><td class="field-name">预答辩意见书：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=5" target="_blank">点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,0,4)">撤销</a></td></tr><%
	End If
	If Not IsNull(rs("TBL_THESIS_FILE3")) Then %>
			<tr><td class="field-name">预答辩论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=6" target="_blank">点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,0,5)">撤销</a></td></tr><%
	End If
	If review_status>=rsDetectPaperUploaded Then
		If Len(thesis_file) Then %>
			<tr><td class="field-name">送检论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=8" target="_blank">点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,0,6)">撤销</a></td></tr><%
		End If %>
			<tr><td class="field-name">送检论文检测报告：</td><td><%
		If IsNull(rs("DETECT_REPORT")) Then
			%>未上传<%
		Else
			%><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=13" target="_blank">点击下载</a><%
		End If
			%>&emsp;<input type="file" name="detect_report" size="30" /></td></tr><%
	End If
	If False And review_status>=rsInstructReviewPaperUploaded Then %>
			<tr><td class="field-name">教指委盲评论文检测报告：</td><td><%
		If IsNull(rs("INSTRUCT_REVIEW_DETECT_REPORT")) Then
			%>未上传<%
		Else
			%><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=14" target="_blank">点击下载</a><%
		End If
			%>&emsp;<input type="file" name="instruct_review_detect_report" size="30" /></td></tr><%
	End If
	If review_status>=rsDetectPaperUploaded And Len(thesis_file_review) Then %>
			<tr><td class="field-name">送审论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=9" target="_blank">点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,0,7)">撤销</a></td></tr><%
	End If
	If review_status>=rsAgreedDetect Then %>
			<tr><td class="field-name">送审申请表：</td><td><%
		If Not IsNull(review_app) Then
			%><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=15" target="_blank" >点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,2,2)">撤销</a>&emsp;<input type="button" id="genReviewApp" value="重新生成送审申请表" /><%
		Else
			%>无&emsp;<input type="button" id="genReviewApp" value="生成送审申请表" /><%
		End If
			%></td></tr><%
	End If
	If review_status>=rsDefencePaperUploaded And Len(thesis_file_modified) Then %>
			<tr><td class="field-name">答辩论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=10" target="_blank">点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,0,8)">撤销</a></td></tr><%
	End If
	If review_status>=rsInstructReviewPaperUploaded And Len(thesis_file_instruct_review) Then %>
			<tr><td class="field-name">教指委盲评论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=11" target="_blank">点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,0,9)">撤销</a></td></tr><%
	End If
	If review_status>=rsFinalPaperUploaded And Len(thesis_file_final) Then %>
			<tr><td class="field-name">定稿论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=12" target="_blank">点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,0,10)">撤销</a></td></tr><%
	End If
	If task_progress>=tpTbl4Uploaded And Len(table_file(4)) Then %>
			<tr><td class="field-name">答辩审批材料：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=7" target="_blank">点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,0,11)">撤销</a></td></tr><%
	End If
			%><tr><td height="10"></td></tr><%
	If review_status>=rsMatchedInstructMember Then %>
			<tr><td class="field-name">当前匹配教指委委员：</td><td>(1)<a href="/index/teacher_resume.asp?id=<%=rs("INSTRUCT_MEMBER1")%>" target="_blank"><%=rs("INSTRUCT_MEMBER_NAME1")%></a>&emsp;(2)<a href="/index/teacher_resume.asp?id=<%=rs("INSTRUCT_MEMBER2")%>" target="_blank"><%=rs("INSTRUCT_MEMBER_NAME2")%></a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,3,4)">撤销</a></td></tr><%
	End If %>
			</td></tr></table>
		</fieldset>
		<fieldset>
			<legend>查重检测记录（按时间先后顺序）<a href="#" onclick="return rollback(<%=paper_id%>,3,0)">撤销全部检测记录</a></legend><%
	If rsDetect.EOF Then
		%>无<%
	Else
		%><ol><%
		Dim index:index=1
		Dim detect_id,detect_time,detect_result,detect_type
		Dim arrDetectStuffNames:arrDetectStuffNames=Array("", "送检论文", "教指委盲评论文")
		Do While Not rsDetect.EOF
			detect_id=rsDetect("Id")
			detect_time=rsDetect("DetectTime")
			If IsNull(detect_time) Then detect_time="无"
			detect_result=rsDetect("Result")
			detect_type=rsDetect("DetectType")
			detect_stuff_name=arrDetectStuffNames(detect_type)
			If IsNull(detect_result) Then detect_result="无" Else detect_result=toNumericString(detect_result)&"%"
			detect_report=rsDetect("ReportFile")
		%><li>类型：<%=detect_stuff_name%>／检测时间：<%=detect_time%>／检测结果：<%=detect_result%>
		<br /><a class="resc" href="fetchDocument.asp?store=detect&type=8&id=<%=detect_id%>" target="_blank"><%=detect_stuff_name%></a><%
			If Not IsNull(detect_report) Then %>
		&emsp;<a class="resc" href="fetchDocument.asp?store=detect&type=13&id=<%=detect_id%>" target="_blank">检测报告</a>
		&emsp;<a href="#" onclick="return deleteDetectResult(<%=paper_id%>,'<%=detect_id%>',0)">删除报告</a><%
			End If
		%>&emsp;<a href="#" onclick="return deleteDetectResult(<%=paper_id%>,'<%=detect_id%>',1)">删除检测记录</a></li><%
			index=index+1
			rsDetect.MoveNext()
		Loop
		%></ol><%
	End If
		%></fieldset>
		<fieldset>
			<legend>送审评阅记录（按时间先后顺序）</legend>
			<table class="form">
			<tr><td class="field-name">当前匹配评阅专家：</td><td>(1)<a href="/index/teacher_resume.asp?id=<%=rs("REVIEWER1")%>" target="_blank"><%=rs("EXPERT_NAME1")%></a>&emsp;(2)<a href="/index/teacher_resume.asp?id=<%=rs("REVIEWER2")%>" target="_blank"><%=rs("EXPERT_NAME2")%></a>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,3,1)">撤销</a></td></tr>
			<tr><td class="field-name">评审结果&nbsp;1：</td><td><%=reviewResultList("review_result",review_result(0),false)%>&emsp;<span class="tip">(A→同意答辩；B→需做适当修改；C→需做重大修改；D→不同意答辩；E→尚未返回)</span></td></tr>
			<tr><td class="field-name">评审结果&nbsp;2：</td><td><%=reviewResultList("review_result",review_result(1),false)%></td></tr>
			<tr><td class="field-name">处理意见：</td><td><%=finalResultList("review_result",review_result(2),false)%></td></tr>
			</table>
			<table class="form">
			<tr><td><table id="datagrid_review_records"></table></td></tr>
			</table>
		</fieldset>
		<fieldset>
			<legend>审核意见记录</legend>
			<table class="form">
				<tr><td><table id="datagrid_audit_records"></table></td></tr>
			</table>
			<table class="form"><%
	If Not IsNull(rs("DETECT_APP_EVAL")) Then %>
			<tr><td class="field-name">导师送检意见：<br /><a href="#" onclick="return rollback(<%=paper_id%>,2,1)">撤销</a></td><td><%=toPlainString(isNullString(rs("DETECT_APP_EVAL"),"未填写"))%></td></tr><%
	End If
	If Not IsNull(rs("REVIEW_APP_EVAL")) Then %>
			<tr><td class="field-name">导师送审意见：<br /><a href="#" onclick="return rollback(<%=paper_id%>,2,2)">撤销</a></td><td>提交时间：<input type="text" class="txt full-width" name="new_submit_review_time" value="<%=rs("SUBMIT_REVIEW_TIME")%>" />
		<br /><%=toPlainString(isNullString(rs("REVIEW_APP_EVAL"),"未填写"))%></td></tr><%
	End If
	If Not IsNull(rs("TUTOR_MODIFY_EVAL")) Then %>
			<tr><td class="field-name"><%=tutor_modify_eval_title%>：<br /><a href="#" onclick="return rollback(<%=paper_id%>,2,3)">撤销</a></td><td><%=toPlainString(isNullString(rs("TUTOR_MODIFY_EVAL"),"未填写"))%></td></tr><%
	End If
	If Not IsNull(rs("DEFENCE_EVAL")) Then %>
			<tr><td class="field-name">答辩委员会修改意见：<br /><a href="#" onclick="return rollback(<%=paper_id%>,3,3)">撤销</a></td><td><%=toPlainString(isNullString(rs("DEFENCE_EVAL"),"未填写"))%></td></tr><%
	End If
	If review_status>=rsMatchedInstructMember Then
		audit_info=getAuditInfo(paper_id,rs("THESIS_FILE4"),auditTypeInstructReview)
			%>
			<tr><td class="field-name">教指委修改意见：</td>
			<td><%
		If IsEmpty(audit_info(0)("AuditorName")) Then
			%>无<%
		Else
			%><ol><%
			For i=0 To UBound(audit_info) %>
			<li>委员：<%=audit_info(i)("AuditorName")%>／审核时间：<%=audit_info(i)("AuditTime")%>&emsp;<a href="#" onclick="return rollback(<%=paper_id%>,3,<%=5+i%>)">撤销</a>
			<br /><%=toPlainString(isNullString(audit_info(i)("Comment"),"未填写"))%></li><%
			Next %>
			</ol><%
		End If
			%></td></tr><%
	End If
	If Not IsNull(rs("DEGREE_MODIFY_EVAL")) Then %>
			<tr><td class="field-name">学院学位评定分会修改意见：<br /><a href="#" onclick="return rollback(<%=paper_id%>,3,7)">撤销</a></td><td><%=toPlainString(isNullString(rs("DEGREE_MODIFY_EVAL"),"未填写"))%></td></tr><%
	End If %>
			</table>
		</fieldset>
		<fieldset>
			<legend>答辩信息<a href="#" onclick="return rollback(<%=paper_id%>,3,2)">撤销</a></legend>
			<table class="form"><%
	If has_defence_plan Then %>
				<tr><td colspan="2"><table id="datagrid_defence_plan"></table></td></tr><%
	End If %>
				<tr><td class="field-name">答辩成绩：</td><td><%=defenceResultList("new_defence_result",defence_result)%></td></tr>
				<tr><td class="field-name">答辩表决结果：</td><td><%=grantDegreeList("new_grant_degree_result",grant_degree_result)%></td></tr>
			</table>
		</fieldset>
		<fieldset>
			<legend>论文状态</legend>
			<table class="form">
				<tr><td class="field-name">表格审核状态：</td><td><select name="new_task_progress"><%
				GetMenuListPubTerm "ReviewStatuses","STATUS_ID1","STATUS_NAME",task_progress,"AND STATUS_ID1 IS NOT NULL"
			%></select></td></tr>
				<tr><td class="field-name">论文审核状态：</td><td><select name="new_review_status"><%
				GetMenuListPubTerm "ReviewStatuses","STATUS_ID2","STATUS_NAME",review_status,"AND STATUS_ID2 IS NOT NULL"
			%></select></td></tr>
			</table>
		</fieldset>
		<table class="form buttons"><tr><td><%
	Select Case opr
		Case 1,2,3,4 %>
		<input type="button" id="reject" name="btnAudit" value="审核不通过<%=arrTable(opr)%>" />
		<input type="button" id="agree" name="btnAudit" value="审核通过<%=arrTable(opr)%>" /><%
		Case 5 %>
		<input type="button" id="reject" name="btnAudit" value="不同意该生论文查重、送审" />
		<input type="button" id="agree" name="btnAudit" value="同意该生论文查重、查重结果低于10%系统自动匹配送审" /><%
		Case 7 %>
		<input type="button" id="audit" name="btnAudit" value="确认评阅结果" /><%
		Case 8 %>
		<input type="button" id="reject" name="btnAudit" value="不同意论文修改" />
		<input type="button" id="agree" name="btnAudit" value="确认修改，同意答辩" /><%
		Case 9 %>
		<input type="button" id="reject" name="btnAudit" value="不同意查重" />
		<input type="button" id="agree" name="btnAudit" value="审核通过，同意查重" /><%
		Case 10 %>
		<input type="button" id="audit1" name="btnAudit" value="以专家一【<%=rs("EXPERT_NAME1")%>】身份评阅该论文" />
		<input type="button" id="audit2" name="btnAudit" value="以专家二【<%=rs("EXPERT_NAME2")%>】身份评阅该论文" /><%
	End Select
	If review_status=rsMatchedReviewer Then
		%><input type="button" value="通知专家评阅" onclick="submitForm(this.form,'notifyReviewer.asp?tid=<%=rs("ID")%>')" /><%
	End If
	If review_status>=rsMatchedReviewer Then
		matchExpertOprName="重新匹配送检论文评阅专家"
	ElseIf review_status=rsAgreedReview Then
		matchExpertOprName="匹配送检论文评阅专家"
	End If
	If review_status>=rsMatchedInstructMember Then
		matchInstructMemberOprName="重新匹配教指委委员"
	ElseIf review_status=rsAgreedInstructReview Then
		matchInstructMemberOprName="匹配教指委委员"
	End If
	If Len(matchExpertOprName) Then
		%><input type="button" value="<%=matchExpertOprName%>" onclick="submitForm(this.form,'matchReviewer.asp?tid=<%=paper_id%>')" /><%
	End If
	If Len(matchInstructMemberOprName) Then
		%><input type="button" value="<%=matchInstructMemberOprName%>" onclick="submitForm(this.form,'matchInstructMember.asp?tid=<%=paper_id%>')" /><%
	End If %>
		<input type="submit" value="确 定" />
		<input type="button" value="关 闭" onclick="closeWindow()" />
		</td></tr></table>
		<input type="hidden" name="opr" value="<%=opr%>" />
		<input type="hidden" id="submittype" name="submittype" />
		<input type="hidden" name="In_ActivityId2" value="<%=activity_id%>">
		<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
		<input type="hidden" name="In_CLASS_ID2" value="<%=class_id%>" />
		<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
		<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
		<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
		<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
		<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
		<input type="hidden" name="pageNo2" value="<%=pageNo%>" />
	</form><%
	If Len(notice_text) Then %>
	<table class="form notice-text"><tr><td><p>论文检测结果及论文评审结果说明：</p>
	<%=notice_text%>
	</td></tr></table><%
	End If %>
</div>
<form id="ret" name="ret" action="paperList.asp" method="post">
<input type="hidden" name="In_ActivityId" value="<%=activity_id%>">
<input type="hidden" name="In_TEACHTYPE_ID" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize" value="<%=pageSize%>" />
<input type="hidden" name="pageNo" value="<%=pageNo%>" /></form>
<script type="text/javascript">
	$(function() {
		$("#activity_id").combobox({
			url: "../api/get-activities",
			valueField: "id",
			textField: "name",
			width: 400,
			editable: false,
			loadFilter: Common.curryLoadFilter(Array.prototype.reverse),
			onLoadFailed: Common.curryOnLoadFailed("获取评阅活动列表"),
			onLoadSuccess: function() {
				$(this).combobox("select", <%=activity_id%>);
			}
		});
		initAuditRecordsDataGrid($("#datagrid_audit_records"), <%=paper_id%>);
		initReviewRecordsDataGrid($("#datagrid_review_records"), <%=paper_id%>, true);
		<%
	If has_defence_plan Then
		Dim defence_time,defence_place,defence_members,defence_memo
		defence_time=toJsString(toPlainString(rs("DEFENCE_TIME")))
		defence_place=toJsString(toPlainString(rs("DEFENCE_PLACE")))
		defence_members=toJsString(toPlainString(rs("DEFENCE_MEMBER")))
		defence_memo=toJsString(toPlainString(rs("MEMO")))
		If IsNull(defence_memo) Then defence_memo="-" %>
		initDefencePlanDataGrid($("#datagrid_defence_plan"),
			"<%=defence_time%>",
			"<%=defence_place%>",
			"<%=defence_members%>",
			"<%=defence_memo%>"
		);<%
	End If %>
		$.each($(":button[name='btnAudit']"),function(index, btn) {
			$(btn).attr({action:["<%=actionUrl1%>","<%=actionUrl2%>"][index],disabled:false}).click(function() {
				$(this).val("正在提交，请稍候……").attr("disabled",true);
				this.form.submittype.value=this.id;
				this.form.action=$(this).attr("action");
				if(!this.form.action.match(/updatePaper\.asp/)){
					this.form.encoding='';
				}
				this.form.submit();
			});
		});
		$(":button#genReviewApp").click(function() {
			$(this.form).attr({action:"genReviewApp.asp?tid=<%=paper_id%>",encoding:''}).submit();
		});
<%
	If review_status=rsAgreedDetect Or review_status=rsDetectUnpassed Or review_status=rsRedetectUnpassed Or review_status=rsAgreedReview Then
		Dim new_review_status_passed:new_review_status_passed=rsAgreedReview
		Dim new_review_status_unpassed
		
		If detect_count>1 Then
			new_review_status_unpassed=rsRedetectUnpassed
		Else
			new_review_status_unpassed=rsDetectUnpassed
		End If
%>
		$(":input[name='reproduct_ratio']").change(function() {
			if(isNaN(this)) return;
			$("select[name='new_review_status']").val(
				!this.trim().length?<%=rsAgreedDetect%>:
				parseFloat(this)<=10?<%=new_review_status_passed%>:<%=new_review_status_unpassed%>
			);
		});<%
	End If
%>
	});
</script></body></html><%
	CloseRs rsRevType
Case 2	' 填写评语页面
	opr=Request.Form("opr")
	submittype=Request.Form("submittype")
	is_reject=submittype="reject"
	Select Case opr
	Case 1,2,3,4
		If is_reject Then
			operation_name="您审核不通过"&arrTable(opr)&"，请填写审核意见"
		ElseIf opr=4 Then
			operation_name="您审核通过了"&arrTable(opr)&"，请填写指导教师意见"
		End If
	Case 5
		If is_reject Then
			operation_name="您不同意论文检测和送审，请填写审核意见"
		Else
			operation_name="您同意了论文检测和送审，请填写送审评语"
		End If
	Case 6
		If is_reject Then
			operation_name="您不同意论文送审，请填写审核意见"
		Else
			operation_name="您同意了论文送审，请填写送审评语"
		End If
	Case 8
		operation_name="填写修改意见"
	Case 9
		operation_name="填写教指委盲评论文审核意见"
	End Select
	tutor_duty_name=getProDutyNameOf(tutor_id)
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>填写审核意见</title>
<% useStylesheet "admin" %>
<% useScript "jquery", "common", "paper" %>
<style type="text/css">
	input[type="text"] { background:none;border-top:0;border-left:0;border-right:0;border-bottom:1px solid dimgray }
</style>
</head>
<body>
<h1><%=operation_name%></h1>
<div class="box">
	<form id="fmOper" action="updatePaper.asp?tid=<%=paper_id%>" method="post" enctype="multipart/form-data" style="margin-top:0;padding-top:10px">
		<table class="form">
		<tr><td>
			<div class="fields">
				<div>作者姓名：<input type="text" class="txt full-width" name="author" value="<%=rs("STU_NAME")%>" readonly /></div>
				<div>学号：<input type="text" class="txt full-width" name="stuno" value="<%=rs("STU_NO")%>" readonly /></div>
				<div>导师姓名、职称：<input type="text" class="txt full-width" name="tutorinfo" value="<%=rs("TUTOR_NAME")%>&nbsp;<%=tutor_duty_name%>" readonly /></div>
			</div>
			<div class="fields">
				学院名称：<input type="text" class="txt full-width" name="faculty" value="工商管理学院" readonly />
			</div>
			<div class="fields">
				学位论文题目：<input type="text" class="txt full-width" name="subject" size="70" value="<%=rs("THESIS_SUBJECT")%>" />
			</div><%
			Select Case opr
			Case 1,2,3,4 ' 填写表格审核意见页面
				Select Case opr
				Case 1
					If Not IsNull(rs("TABLE_FILE1")) Then %>
			<div class="fields">开题报告表：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=1" target="_blank">点击下载</a></div><%
					End If
					If Not IsNull(rs("TBL_THESIS_FILE1")) Then %>
			<div class="fields">开题论文：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=2" target="_blank">点击下载</a></div><%
					End If
				Case 2
					If Not IsNull(rs("TABLE_FILE2")) Then %>
			<div class="fields">中期考核表：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=3" target="_blank">点击下载</a></div><%
					End If
					If Not IsNull(rs("TBL_THESIS_FILE2")) Then %>
			<div class="fields">中期论文：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=4" target="_blank">点击下载</a></div><%
					End If
				Case 3
					If Not IsNull(rs("TABLE_FILE3")) Then %>
			<div class="fields">预答辩意见书：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=5" target="_blank">点击下载</a></div><%
					End If
					If Not IsNull(rs("TBL_THESIS_FILE3")) Then %>
			<div class="fields">预答辩论文：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=6" target="_blank">点击下载</a></div><%
					End If
				Case 4 %>
			<div class="fields">答辩审批材料：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=7" target="_blank">点击下载</a></div><%
				End Select
				If is_reject Then
					comment_name=arrTable(opr)&"审核意见（200-2000字）："
				ElseIf opr=4 Then
					comment_name="校内指导教师意见（包括对申请人的学习情况、思想表现及论文的学术评语，科研工作能力和完成科研工作情况，以及是否同意申请学位论文答辩的意见，200-2000字）"
				End If %>
			<div class="fields"><%=comment_name%></div>
			<div class="fields"><textarea name="comment" rows="15" style="width:100%"><%=comment%></textarea></div>
			<div class="fields"><span id="comment_tip"></span></div><%
				Case 5 ' 填写论文送检送审意见页面 %>
			<div class="fields">送检论文：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=8" target="_blank">点击下载</a>&emsp;送审论文：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=9" target="_blank">点击下载</a></div>
			<div class="fields">导师对论文的意见<span class="comment-notice">（200-2000字，包含选题意义；文献资料的掌握；数据、材料的收集、论证、结论是否合理；基本论点、结论和建议有无理论意义和实践价值等）</span>：</div>
			<div class="fields"><span class="tip">提示：复制比低于10%的学员，系统会自动匹配进行论文盲审。复制比高于10%的学员，由中心统一进行二次查重，二次查重所产生的费用由学员本人缴纳。</span></div>
			<div class="fields"><textarea name="comment" rows="15" style="width:100%"><%=comment%></textarea></div>
			<div class="fields"><span id="comment_tip"></span></div><%
				Case 6 ' 填写导师送审评语页面 %>
			<div class="fields">送审论文：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=9" target="_blank">点击下载</a></div>
			<div class="fields">导师对学位论文的评语<span class="comment-notice">（请阅读论文后填写，200-2000字）</span>：</div>
			<div class="fields">送审评语的基本内容参考：<br /><%=getNoticeText(rs("TEACHTYPE_ID"),"review_eval_reference")%></div>
			<div class="fields"><textarea name="comment" rows="10" style="width:100%"><%=comment%></textarea></div>
			<div class="fields"><span id="comment_tip"></span></div><%
					If Not is_reject Then %>
			<div class="fields">
			<table class="template">
			<tr><td width="100" align="center">作者承诺</td>
			<td><p>&emsp;&emsp;1．该学位论文为公开学位论文，其中不涉及国家秘密项目和其它不宜公开的内容，否则将由本人承担因学位论文涉密造成的损失和相关的法律责任；<br />
			&emsp;&emsp;2．该学位论文是本人在导师的指导下独立进行研究所取得的研究成果，不存在学术不端行为。</p>
			<p align="right">作者签名：&emsp;&emsp;&emsp;&emsp;&nbsp;日期：<%=FormatDateTime(Now(),1)%></p></td></tr>
			<tr><td align="center">指导教师<br />意见</td>
			<td><p><span style="font-size:15pt">■</span>&nbsp;同意送审<br /><span style="font-size:15pt">□</span>&nbsp;不同意送审</p>
			<p align="right">指导教师签名：&emsp;&emsp;&emsp;&emsp;&nbsp;日期：<span style="visibility:hidden"><%=FormatDateTime(Now(),1)%></span></p></td></tr>
			<tr><td align="center"></td>
			<td><p><span style="font-size:15pt">□</span>&nbsp;同意送审<br /><span style="font-size:15pt">□</span>&nbsp;不同意送审</p>
			<p align="right">主管院领导签名：&emsp;&emsp;&emsp;&emsp;&nbsp;日期：<span style="visibility:hidden"><%=FormatDateTime(Now(),1)%></span></p></td></tr>
			<tr><td align="center">备注</td>
			<td><p>经图书馆检测，学位论文文字复制比&nbsp;<span style="text-decoration:underline"><%=toNumericString(reproduct_ratio)%>%</span><input type="hidden" name="reproduct_ratio" size="10" value="<%=reproduct_ratio%>" /></p></td></tr>
			</table>
			</div><%
					End If
			End Select %>
		</td></tr></table>
		<table class="form buttons"><tr><td>
		<input type="button" name="btnSubmit" value="提 交" />
		<input type="button" value="返 回" onclick="history.go(-1)" />
		<input type="button" value="关 闭" onclick="closeWindow()" />
		</td></tr></table>
		<input type="hidden" name="opr" value="<%=opr%>" />
		<input type="hidden" name="submittype" value="<%=submittype%>" />
		<input type="hidden" name="In_ActivityId2" value="<%=activity_id%>">
		<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
		<input type="hidden" name="In_CLASS_ID2" value="<%=class_id%>" />
		<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
		<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
		<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
		<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
		<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
		<input type="hidden" name="pageNo2" value="<%=pageNo%>" />
	</form>
</div>
<form id="ret" name="ret" action="paperList.asp" method="post">
<input type="hidden" name="In_ActivityId" value="<%=activity_id%>">
<input type="hidden" name="In_TEACHTYPE_ID" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize" value="<%=pageSize%>" />
<input type="hidden" name="pageNo" value="<%=pageNo%>" /></form>
<script type="text/javascript">
	$(":input[name='comment']").on("input propertychange",function(){checkLength(this,2000);});
	$(":button[name='btnSubmit']").attr("disabled",false).click(function() {
		if(confirm("确定要提交吗？")) {
			$(this).val("正在提交，请稍候……").attr("disabled",true);
			this.form.submit();
		}
	});
</script></body></html><%
End Select
CloseRs rsDetect
CloseRs rs
CloseConn conn
%>
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")
If Not hasPrivilege(Session("Treadprivileges"),"I11") Then Response.Redirect("../error.asp?no-privilege")
step=Request.QueryString("step")
paper_id=Request.QueryString("tid")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
spec_id=Request.Form("In_SPECIALITY_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
teacher_id=Session("TId")
If Len(paper_id)=0 Or Not IsNumeric(paper_id) Then
	showErrorPage "参数无效。", "提示"
End If

Connect conn
sql="SELECT * FROM ViewDissertations_tutor WHERE ID="&paper_id
GetRecordSet conn,rs,sql,count
If count=0 Then
	CloseRs rs
	CloseConn conn
	showErrorPage "数据库没有该论文记录！", "提示"
End If

Dim section_access_info
Set section_access_info=getSectionAccessibilityInfo(rs("ActivityId"),rs("TEACHTYPE_ID"),sectionAudit)
Dim arrCaptionAgreed:arrCaptionAgreed=Array("","审核通过开题报告表，同意参加开题报告","审核通过中期考核表","审核通过预答辩意见书，同意参加预答辩","审核通过审批材料，同意参加答辩")
Dim arrCaptionRefused:arrCaptionRefused=Array("","审核不通过，不同意参加开题报告（延期3-6个月重新做开题报告）","审核不通过中期考核表","审核不通过预答辩意见书，不同意参加预答辩","审核不通过审批材料，不同意参加答辩")
Dim review_status,review_result(2),reviewer_master_level(1),review_file(1),review_time(1),review_level(1)

sql="SELECT * FROM DetectResults WHERE DissertationId="&paper_id
GetRecordSet conn,rsDetect,sql,count
If rs("REVIEWER1")=teacher_id Then
	reviewer=0
ElseIf rs("REVIEWER2")=teacher_id Then
	reviewer=1
End If
If rs("INSTRUCT_MEMBER1")=teacher_id Then
	instruct_member=0
ElseIf rs("INSTRUCT_MEMBER2")=teacher_id Then
	instruct_member=1
End If
If rs("TEACHTYPE_ID")=5 Then
	reviewfile_type=2
Else
	reviewfile_type=1
End If
tutor_id=rs("TUTOR_ID")
review_app=rs("REVIEW_APP")
task_progress=rs("TASK_PROGRESS")
review_status=rs("REVIEW_STATUS")
stat_text1=rs("STAT_TEXT1")
stat_text2=rs("STAT_TEXT2")
is_review_visible=Array(rs("ReviewFileDisplayStatus1") >= 1,rs("ReviewFileDisplayStatus2") >= 1)
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
Case rsMatchedInstructMember:opr=11
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
Select Case step
Case vbNullString	' 论文详情页面

	updateActiveTime teacher_id
	
	Dim tutor_modify_eval_title
	arrActionUrlList=Array("?tid="&paper_id&"&step=2","updatePaper.asp?tid="&paper_id,"../expert/paperDetail.asp?tid="&paper_id&"&step=2")
	Select Case opr
	Case 1,2,3
		actionUrl1=arrActionUrlList(0)
		actionUrl2=arrActionUrlList(1)
	Case 4,5,6,8,9,11
		actionUrl1=arrActionUrlList(0)
		actionUrl2=actionUrl1
	Case 7
		actionUrl1=arrActionUrlList(1)
		actionUrl2=actionUrl1
	Case 10
		actionUrl1=arrActionUrlList(2)
		actionUrl2=vbNullString
	End Select
	If review_status>=rsAgreedDefence Then
		tutor_modify_eval_title="导师同意答辩意见"
	ElseIf review_status=rsRefusedDefence Then
		tutor_modify_eval_title="导师不同意答辩意见"
	Else
		tutor_modify_eval_title="导师对答辩论文的意见"
	End If

	notice_text = getNoticeText(rs("TEACHTYPE_ID"), "review_result_desc")
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>查看论文信息</title>
<% useStylesheet "tutor", "jeasyui" %>
<% useScript "jquery", "jeasyui", "common", "*paper" %>
</head>
<body>
<h1>专业硕士论文详情<h2>论文当前状态：【<%=stat%>】</h2></h1><%
	If opr<>0 And Not section_access_info("accessible") Then
%><p align="center"><span class="tip"><%=section_access_info("tip")%></span></p><%
	End If %>
<div class="box">
	<form id="fmDetail" action="<%=actionUrl%>" method="post">
		<fieldset>
			<legend>论文基本情况</legend>
			<table class="form">
			<tr><td class="field-name">评阅活动：</td><td><input type="text" class="txt full-width" size="95%" value="<%=rs("ActivityName")%>" readonly /></td></tr>
			<tr><td class="field-name">论文题目：</td><td><input type="text" class="txt full-width" name="subject" size="95%" value="<%=rs("THESIS_SUBJECT")%>" readonly /></td></tr>
			<tr><td class="field-name">（英文）：</td><td><input type="text" class="txt full-width" name="subject_en" size="85%" value="<%=rs("THESIS_SUBJECT_EN")%>" readonly /></td></tr>
			<tr><td class="field-name">作者姓名：</td><td><%=rs("STU_NAME")%><a class="open-window" href="#" onclick="return showStudentProfile(<%=rs("STU_ID")%>,2)">查看学生资料</a></td></tr>
			<tr><td class="field-name">学号：</td><td><input type="text" class="txt full-width" name="stuno" size="15" value="<%=rs("STU_NO")%>" readonly /></td></tr>
			<tr><td class="field-name">学位类别：</td><td><input type="text" class="txt full-width" name="degreename" size="10" value="<%=rs("TEACHTYPE_NAME")%>" readonly /></td></tr>
			<tr><td class="field-name">指导教师：</td><td><input type="text" class="txt full-width" name="tutorname" size="95%" value="<%=rs("TUTOR_NAME")%>" readonly /></td></tr><%
	If reviewfile_type=2 Then %>
			<tr><td class="field-name">领域名称：</td><td><input type="text" class="txt full-width" name="speciality" size="95%" value="<%=rs("SPECIALITY_NAME")%>" readonly /></td></tr><%
	End If %>
			<tr><td class="field-name">研究方向：</td><td><input type="text" class="txt full-width" name="researchway_name" size="95%" value="<%=rs("RESEARCHWAY_NAME")%>" readonly /></td></tr>
			<tr><td class="field-name">论文关键词：</td><td><input type="text" class="txt full-width" name="keywords_ch" size="85%" value="<%=rs("KEYWORDS")%>" readonly /></td></tr>
			<tr><td class="field-name">（英文）：</td><td><input type="text" class="txt full-width" name="keywords_en" size="85%" value="<%=rs("KEYWORDS_EN")%>" readonly /></td></tr>
			<tr><td class="field-name">院系名称：</td><td><input type="text" class="txt full-width" name="faculty" size="30%" value="工商管理学院" readonly /></td></tr>
			<tr><td class="field-name">班级：</td><td><input type="text" class="txt full-width" name="class" size="51%" value="<%=rs("CLASS_NAME")%>" readonly /></td></tr><%
	If Not IsNull(rs("THESIS_FORM")) And Len(rs("THESIS_FORM")) Then %>
			<tr><td class="field-name">论文形式：</td><td><input type="text" class="txt full-width" name="thesisform" size="95%" value="<%=rs("THESIS_FORM")%>" readonly /></td></tr><%
	End If
	If review_status>=rsDetectUnpassed Then %>
			<tr><td class="field-name">送检论文文字复制比：</td><td><input type="text" class="txt full-width" name="reproduct_ratio" size="10px" value="<%=toNumericString(reproduct_ratio)%>%" readonly /></td></tr><%
	End If
	If False And review_status>=rsInstructReviewPaperDetected Then %>
			<tr><td class="field-name">教指委盲评论文文字复制比：</td><td><input type="text" class="txt full-width" name="instruct_review_reproduct_ratio" size="10px" value="<%=toNumericString(instruct_review_reproduct_ratio)%>%" readonly /></td></tr><%
	End If
	If task_progress>=tpTbl1Uploaded Then
		If Len(table_file(1)) Then %>
			<tr><td class="field-name">开题报告表：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=1" target="_blank">点击下载</a></td></tr><%
		End If
		If Not IsNull(rs("TBL_THESIS_FILE1")) Then %>
			<tr><td class="field-name">开题论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=2" target="_blank">点击下载</a></td></tr><%
		End If
	End If
	If task_progress>=tpTbl2Uploaded Then
		If Len(table_file(2)) Then %>
			<tr><td class="field-name">中期考核表：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=3" target="_blank">点击下载</a></td></tr><%
		End If
		If Not IsNull(rs("TBL_THESIS_FILE2")) Then %>
			<tr><td class="field-name">中期论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=4" target="_blank">点击下载</a></td></tr><%
		End If
	End If
	If task_progress>=tpTbl3Uploaded Then
		If Len(table_file(3)) Then %>
			<tr><td class="field-name">预答辩意见书：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=5" target="_blank">点击下载</a></td></tr><%
		End If
		If Not IsNull(rs("TBL_THESIS_FILE3")) Then %>
			<tr><td class="field-name">预答辩论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=6" target="_blank">点击下载</a></td></tr><%
		End If
	End If
	If review_status>=rsDetectPaperUploaded Then
		If Len(thesis_file) Then %>
			<tr><td class="field-name">送检论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=8" target="_blank">点击下载</a></td></tr><%
		End If %>
			<tr><td class="field-name">送检论文检测报告：</td><td><%
		If IsNull(rs("DETECT_REPORT")) Then
			%>未上传<%
		Else
			%><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=13" target="_blank">点击下载</a><%
		End If
			%></td></tr><%
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
			<tr><td class="field-name">送审论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=9" target="_blank">点击下载</a></td></tr><%
	End If
	If review_status>=rsAgreedReview Then
		If Not IsNull(review_app) Then %>
			<tr><td class="field-name">送审申请表：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=15" target="_blank" >点击下载</a></td></tr><%
		End If
	End If
	If review_status>=rsDefencePaperUploaded And Len(thesis_file_modified) Then %>
			<tr><td class="field-name">答辩论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=10" target="_blank">点击下载</a></td></tr><%
	End If
	If review_status>=rsInstructReviewPaperUploaded And Len(thesis_file_instruct_review) Then %>
			<tr><td class="field-name">教指委盲评论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=11" target="_blank">点击下载</a></td></tr><%
	End If
	If review_status>=rsFinalPaperUploaded And Len(thesis_file_final) Then %>
			<tr><td class="field-name">定稿论文：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=12" target="_blank">点击下载</a></td></tr><%
	End If
	If task_progress>=tpTbl4Uploaded And Len(table_file(4)) Then %>
			<tr><td class="field-name">答辩审批材料：</td><td><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=7" target="_blank">点击下载</a></td></tr><%
	End If
			%>
			</table>
		</fieldset>
		<fieldset>
		<legend>查重检测记录（按时间先后顺序）</legend><%
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
		&emsp;<a class="resc" href="fetchDocument.asp?store=detect&type=13&id=<%=detect_id%>" target="_blank">检测报告</a><%
			End If
			index=index+1
			rsDetect.MoveNext()
		Loop
		%></ol><%
	End If
		%></fieldset>
		<fieldset>
			<legend>送审评阅记录（按时间先后顺序）</legend>
			<table class="form"><%
			' 根据评阅书显示设置决定是否显示文件
			If is_review_visible(0) Then %>
			<tr><td class="field-name">评审结果&nbsp;1：</td><td><%=reviewResultList("review_result",review_result(0),false)%>&emsp;<span class="tip">(A→同意答辩；B→需做适当修改；C→需做重大修改；D→不同意答辩；E→尚未返回)</span></td></tr><%
			End If
			If is_review_visible(1) Then %>
			<tr><td class="field-name">评审结果&nbsp;2：</td><td><%=reviewResultList("review_result",review_result(1),false)%></td></tr><%
			End If
			If is_review_visible(0) And is_review_visible(1) Then %>
			<tr><td class="field-name">处理意见：</td><td><%=finalResultList("review_result",review_result(2),false)%></td></tr><%
			End If %>
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
	If review_status>=rsRefusedDetect Then %>
				<tr><td class="field-name">导师送检意见：</td><td><%=toPlainString(isNullString(rs("DETECT_APP_EVAL"),"未填写"))%></td></tr><%
	End If
	If review_status=rsAgreedDetect Or review_status>=rsAgreedReview Then
		submit_review_time=rs("SUBMIT_REVIEW_TIME")
		If Not IsNull(submit_review_time) Then submit_review_time="("&submit_review_time&")"
			%>
				<tr><td colspan="2"><hr /></td></tr>
				<tr><td class="field-name">导师送审意见：<br /><%=submit_review_time%></td><td><%=toPlainString(isNullString(rs("REVIEW_APP_EVAL"),"未填写"))%></td></tr><%
	End If
	If review_status>=rsRefusedDefence Then %>
				<tr><td colspan="2"><hr /></td></tr>
				<tr><td class="field-name"><%=tutor_modify_eval_title%>：</td><td><%=toPlainString(isNullString(rs("TUTOR_MODIFY_EVAL"),"未填写"))%></td></tr><%
	End If
	If review_status>=rsDefenceEval Then %>
				<tr><td colspan="2"><hr /></td></tr>
				<tr><td class="field-name">答辩委员会修改意见：</td><td><%=toPlainString(isNullString(rs("DEFENCE_EVAL"),"未填写"))%></td></tr><%
	End If
	If Not IsNull(rs("DEGREE_MODIFY_EVAL")) Then %>
				<tr><td colspan="2"><hr /></td></tr>
				<tr><td class="field-name">学院学位评定分会修改意见：</td><td><%=toPlainString(isNullString(rs("DEGREE_MODIFY_EVAL"),"未填写"))%></td></tr><%
	End If %>
			</table>
		</fieldset>
		<fieldset>
		<legend>答辩信息</legend>
			<table class="form"><%
			If has_defence_plan Then %>
				<tr><td colspan="2"><table id="datagrid_defence_plan"></table></td></tr><%
			End If %>
				<tr><td class="field-name">答辩成绩：</td><td><%=defenceResultList("new_defence_result",defence_result)%></td></tr>
				<tr><td class="field-name">答辩表决结果：</td><td><%=grantDegreeList("new_grant_degree_result",grant_degree_result)%></td></tr>
			</table>
		</fieldset>
		<table class="form buttons"><tr><td><%
	If section_access_info("accessible") Then
		Select Case opr
		Case 1,2,3,4 %>
		<input type="button" id="reject" name="btnsubmit" value="<%=arrCaptionRefused(opr)%>" />&emsp;
		<input type="button" id="agree" name="btnsubmit" value="<%=arrCaptionAgreed(opr)%>" />&emsp;<%
		Case 5 %>
		<input type="button" id="reject" name="btnsubmit" value="不同意该生论文查重、送审" />&emsp;
		<input type="button" id="agree" name="btnsubmit" value="同意该生论文查重、查重结果低于10%系统自动匹配送审" />&emsp;<%
		Case 7
			If is_review_visible(0) And is_review_visible(1) Then %>
		<input type="button" name="btnsubmit" value="确认评阅结果" />&emsp;<%
			End If
		Case 8 %>
		<input type="button" id="reject" name="btnsubmit" value="不同意论文修改" />&emsp;
		<input type="button" id="agree" name="btnsubmit" value="确认修改，同意答辩" />&emsp;<%
		Case 9 %>
		<input type="button" id="reject" name="btnsubmit" value="不同意查重" />&emsp;
		<input type="button" id="agree" name="btnsubmit" value="审核通过，同意查重" />&emsp;<%
		Case 10
			If Not IsEmpty(reviewer) Then %>
		<input type="button" name="btnsubmit" value="评阅该论文" />&emsp;<%
			End If
		Case 11
			If Not IsEmpty(instruct_member) Then %>
		<input type="button" id="agree" name="btnsubmit" value="审核该论文" />&emsp;<%
			End If
		End Select
	End If
		%><input type="button" value="关 闭" onclick="closeWindow()" />
		</td></tr></table>
		<input type="hidden" name="opr" value="<%=opr%>" />
		<input type="hidden" id="submittype" name="submittype" />
		<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
		<input type="hidden" name="In_SPECIALITY_ID2" value="<%=spec_id%>" />
		<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
		<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
		<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
		<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
		<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
		<input type="hidden" name="pageNo2" value="<%=pageNo%>" />
	</form>
	<% If Len(notice_text) Then %>
	<table class="form notice-text"><tr><td><p>论文检测结果及论文评审结果说明：</p>
	<%=notice_text%>
	</td></tr></table>
	<% End If %>
</div>
<form id="ret" name="ret" action="paperList.asp" method="post">
<input type="hidden" name="In_TEACHTYPE_ID" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID" value="<%=spec_id%>" />
<input type="hidden" name="In_ENTER_YEAR" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize" value="<%=pageSize%>" />
<input type="hidden" name="pageNo" value="<%=pageNo%>" /></form>
<script type="text/javascript">
	initAuditRecordsDataGrid($("#datagrid_audit_records"), <%=paper_id%>);
	initReviewRecordsDataGrid($("#datagrid_review_records"), <%=paper_id%>, false);
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
	var btnsubmit=document.getElementsByName("btnsubmit");
	var arrActionUrl=["<%=actionUrl1%>","<%=actionUrl2%>"];
	if(btnsubmit) {
		for(var i=0;i<btnsubmit.length;i++) {
			btnsubmit.item(i).action=arrActionUrl[i];
			btnsubmit.item(i).onclick=function() {
				this.value="正在提交，请稍候……";
				this.disabled=true;
				this.form.submittype.value=this.id;
				this.form.action=this.action;
				this.form.submit();
			}
			btnsubmit.item(i).disabled=false;
		}
	}
</script></body></html><%
Case 2	' 填写评语页面
	opr=Request.Form("opr")
	submittype=Request.Form("submittype")
	is_reject=submittype="reject"
	Select Case opr
	Case 1,2,3,4
		If is_reject Then
			operation_name="您审核不通过"&arrTable(opr)&"，请填写审核意见"
		ElseIf opr=4 Then
			operation_name="您审核通过了"&arrTable(opr)&"，请填写指导意见"
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
		If is_reject Then
			operation_name="您不同意论文修改，请填写意见"
		Else
			operation_name="您同意该生参加论文答辩，请填写意见"
		End If
	Case 9
		operation_name="填写教指委盲评论文审核意见"
	Case 11
		operation_name="填写教指委盲评论文修改意见"
	End Select
	tutor_duty_name=getProDutyNameOf(tutor_id)
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>填写审核意见</title>
<% useStylesheet "tutor" %>
<% useScript "jquery", "common", "paper", "*drafting" %>
<style type="text/css">
	input[type="text"] { background:none;border-top:0;border-left:0;border-right:0;border-bottom:1px solid dimgray }
</style>
</head>
<body>
<center><font size=4><b><%=operation_name%></b></font>
<form id="fmOper" action="updatePaper.asp?tid=<%=paper_id%>" method="post" style="margin-top:0;padding-top:10px">
<table class="form" width="800" cellspacing="1" cellpadding="3">
<tr><td>作者姓名：<a href="#" onclick="return showStudentProfile(<%=rs("STU_ID")%>,2)"><%=rs("STU_NAME")%></a></td>
<td>学号：<input type="text" class="txt full-width" name="stuno" value="<%=rs("STU_NO")%>" readonly /></td>
<td>导师姓名、职称：<input type="text" class="txt full-width" name="tutorinfo" value="<%=Session("Tname")%>&nbsp;<%=tutor_duty_name%>" readonly /></td></tr>
<tr><td colspan=2>申请学位专业名称：<input type="text" class="txt full-width" name="speciality" size="50" value="<%=rs("SPECIALITY_NAME")%>" readonly /></td>
<td>学院名称：<input type="text" class="txt full-width" name="faculty" value="工商管理学院" readonly /></td></tr>
<tr><td colspan=3>学位论文题目：<input type="text" class="txt full-width" name="subject" size="70" value="<%=rs("THESIS_SUBJECT")%>" readonly /></td></tr><%
	Select Case opr
	Case 1,2,3,4 ' 填写表格审核意见页面
		Select Case opr
		Case 1
			If Not IsNull(rs("TABLE_FILE1")) Then %>
<tr><td colspan=3>开题报告表：&emsp;&emsp;<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=1" target="_blank">点击下载</a></td></tr><%
			End If
			If Not IsNull(rs("TBL_THESIS_FILE1")) Then %>
<tr><td colspan=3>开题论文：&emsp;&emsp;&emsp;<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=2" target="_blank">点击下载</a></td></tr><%
			End If
		Case 2
			If Not IsNull(rs("TABLE_FILE2")) Then %>
<tr><td colspan=3>中期考核表：&emsp;&emsp;<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=3" target="_blank">点击下载</a></td></tr><%
			End If
			If Not IsNull(rs("TBL_THESIS_FILE2")) Then %>
<tr><td colspan=3>中期论文：&emsp;&emsp;&emsp;<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=4" target="_blank">点击下载</a></td></tr><%
			End If
		Case 3
			If Not IsNull(rs("TABLE_FILE3")) Then %>
<tr><td colspan=3>预答辩意见书：&emsp;<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=5" target="_blank">点击下载</a></td></tr><%
			End If
			If Not IsNull(rs("TBL_THESIS_FILE3")) Then %>
<tr><td colspan=3>预答辩论文：&emsp;&emsp;<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=6" target="_blank">点击下载</a></td></tr><%
			End If
		Case 4 %>
<tr><td colspan=3>答辩审批材料：&emsp;<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=7" target="_blank">点击下载</a></td></tr><%
		End Select
		If is_reject Then
			comment_name=arrTable(opr)&"审核意见（200-2000字）："
		ElseIf opr=4 Then
			comment_name="校内指导教师意见（包括对申请人的学习情况、思想表现及论文的学术评语，科研工作能力和完成科研工作情况，以及是否同意申请学位论文答辩的意见，200-2000字）"
		End If %>
<tr><td colspan=3><%=comment_name%><span id="comment_tip"></span></td></tr>
<tr><td colspan=3><textarea name="comment" rows="15" style="width:100%"></textarea></td></tr><%
	Case 5 ' 填写论文送检送审意见页面
%><tr><td colspan="3">送检论文：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=8" target="_blank">点击下载</a>&emsp;送审论文：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=9" target="_blank">点击下载</a></td></tr><%
		If Not rsDetect.EOF Then
%><tr><td colspan="3">上次论文检测记录：<%
			detect_time=rsDetect("DetectTime")
			If IsNull(detect_time) Then detect_time="无"
			detect_result=rsDetect("Result")
			If IsNull(detect_result) Then detect_result="无" Else detect_result=detect_result&"%"
			detect_report=rsDetect("ReportFile")
%>检测时间：<%=detect_time%>，查重结果：<%=detect_result%><%
			If Not IsNull(detect_report) Then %>
<br/><a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=8&time=<%=detect_time%>" target="_blank">送检论文</a>
<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=13&time=<%=detect_time%>" target="_blank">检测报告</a><%
			End If
%></td></tr><%
		End If
%><tr><td colspan=3>导师对论文的意见<span class="comment-notice">（200-2000字，包含选题意义；文献资料的掌握；数据、材料的收集、论证、结论是否合理；基本论点、结论和建议有无理论意义和实践价值等）</span>：<br/><span class="tip">提示：复制比低于10%的学员，系统会自动匹配进行论文盲审。复制比高于10%的学员，由中心统一进行二次查重，二次查重所产生的费用由学员本人缴纳。</span>&emsp;<span id="comment_tip"></span><br/>
送审评语的基本内容参考：<br/><%=getNoticeText(rs("TEACHTYPE_ID"),"review_eval_reference")%></td></tr>
<tr><td colspan=3><textarea name="comment" rows="15" style="width:100%"></textarea></td></tr><%
	Case 6 ' 填写导师送审评语页面 %>
<tr><td colspan=3>送审论文：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=9" target="_blank">点击下载</a></td></tr>
<tr><td colspan=3>导师对学位论文的评语<span class="comment-notice">（请阅读论文后填写，200-2000字）</span>：<span id="comment_tip"></span><br/>
送审评语的基本内容参考：<br/><%=getNoticeText(rs("TEACHTYPE_ID"),"review_eval_reference")%></td></tr>
<tr><td colspan=3><textarea name="comment" rows="10" style="width:100%"></textarea><br/></td></tr><%
		If Not is_reject Then %>
<tr><td colspan=3 style="padding:0"><table class="form" width="100%" cellspacing="1" cellpadding="3">
<tr><td width="100" align="center">作者承诺</td>
<td><p>&emsp;&emsp;1．该学位论文为公开学位论文，其中不涉及国家秘密项目和其它不宜公开的内容，否则将由本人承担因学位论文涉密造成的损失和相关的法律责任；<br/>
&emsp;&emsp;2．该学位论文是本人在导师的指导下独立进行研究所取得的研究成果，不存在学术不端行为。</p>
<p align="right">作者签名：&emsp;&emsp;&emsp;&emsp;&nbsp;日期：<%=toDateTime(Now(),1)%></p></td></tr>
<tr><td align="center">指导教师<br/>意见</td>
<td><p><span style="font-size:15pt">■</span>&nbsp;同意送审<br/><span style="font-size:15pt">□</span>&nbsp;不同意送审</p>
<p align="right">指导教师签名：&emsp;&emsp;&emsp;&emsp;&nbsp;日期：<span style="visibility:hidden"><%=toDateTime(Now(),1)%></span></p></td></tr>
<tr><td align="center"></td>
<td><p><span style="font-size:15pt">□</span>&nbsp;同意送审<br/><span style="font-size:15pt">□</span>&nbsp;不同意送审</p>
<p align="right">主管院领导签名：&emsp;&emsp;&emsp;&emsp;&nbsp;日期：<span style="visibility:hidden"><%=toDateTime(Now(),1)%></span></p></td></tr>
<tr><td align="center">备注</td>
<td><p>经图书馆检测，学位论文文字复制比&nbsp;<span style="text-decoration:underline"><%=toNumericString(reproduct_ratio)%>%</span><input type="hidden" name="reproduct_ratio" size="10" value="<%=reproduct_ratio%>" /></p></td></tr></table></td></tr><%
		End If
	Case 8 ' 填写修改意见页面 %>
<tr><td colspan=3>答辩论文：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=10" target="_blank">点击下载</a></td></tr>
<tr><td colspan=3>导师对学位论文的评语<span class="comment-notice">（此意见将嵌入学生《学位论文答辩及授予学位审批材料》中，并进入学籍档案，意见需包含对学生的业务学习、思想表现及学位论文的学术评语，科研工作能力和完成科研工作情况，以及是否同意申请学位论文答辩的意见，200-2000字）</span>：<span id="comment_tip"></span></td></tr>
<tr><td colspan=3><textarea name="comment" rows="10" style="width:100%"></textarea><br/></td></tr><%
	Case 9 ' 填写教指委盲评论文审核意见页面 %>
<tr><td colspan=3>教指委盲评论文：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=11" target="_blank">点击下载</a></td></tr>
<tr><td colspan=3>导师对学位论文的评语：<span id="comment_tip"></span></td></tr>
<tr><td colspan=3><textarea name="comment" rows="10" style="width:100%"></textarea><br/></td></tr><%
	Case 11 ' 填写教指委盲评论文修改意见页面 %>
<tr><td colspan=3>教指委盲评论文：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=11" target="_blank">点击下载</a></td></tr>
<tr><td colspan=3>对学位论文的修改意见：<span id="comment_tip"></span></td></tr>
<tr><td colspan=3><textarea name="comment" rows="10" style="width:100%"></textarea><br/></td></tr><%
	End Select
	If Not IsNull(comment) Then %>
<tr><td colspan=3><%=comment%></td></tr><%
	End If %>
<tr class="buttons">
<td colspan="3"><p align="center"><input type="button" id="btnsavedraft" value="保存草稿" />&emsp;
<input type="button" id="btnloaddraft" value="读取草稿" />&emsp;
<input type="button" id="btnsubmit" name="btnsubmit" value="提 交" />&emsp;
<input type="button" value="返 回" onclick="history.go(-1)" />&emsp;
<input type="button" value="关 闭" onclick="closeWindow()" />
</p></td></tr></table>
<input type="hidden" name="opr" value="<%=opr%>" />
<input type="hidden" name="submittype" value="<%=submittype%>" />
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID2" value="<%=spec_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form></center>
<form id="ret" name="ret" action="paperList.asp" method="post">
<input type="hidden" name="In_TEACHTYPE_ID" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID" value="<%=spec_id%>" />
<input type="hidden" name="In_ENTER_YEAR" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize" value="<%=pageSize%>" />
<input type="hidden" name="pageNo" value="<%=pageNo%>" /></form>
<script type="text/javascript">
	$(document).ready(function(){
		$('[name="comment"]').keyup(function(){checkLength(this,2000)});
		if($('#btnsubmit').size()>0) {
			$('#btnsubmit').click(function() {
				if(confirm("确定要提交吗？")) {
					$(this).val("正在提交，请稍候……")
								 .attr('disabled',true);
					this.form.submit();
				}
			}).attr('disabled',false);
		}
		$('#btnsavedraft').click(function() {
			saveAsDraft(<%=paper_id%>);
		});
		verifyDraft(<%=paper_id%>);
		// 每30秒对草稿进行自动保存
		setInterval('saveAsDraft(<%=paper_id%>,true)',30000);
	});
</script></body></html><%
End Select
CloseRs rsDetect
CloseRs rs
CloseConn conn
%>
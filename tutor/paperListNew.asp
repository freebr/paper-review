<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")

Dim PubTerm,PageNo,PageSize
activity_id=toUnsignedInt(Request.Form("In_ActivityId"))
teachtype_id=toUnsignedInt(Request.Form("In_TEACHTYPE_ID"))
spec_id=toUnsignedInt(Request.Form("In_SPECIALITY_ID"))
enter_year=toUnsignedInt(Request.Form("In_ENTER_YEAR"))
query_task_progress=toUnsignedInt(Request.Form("In_TASK_PROGRESS"))
query_review_status=toUnsignedInt(Request.Form("In_REVIEW_STATUS"))
is_instruct_review=Request.QueryString()="instruct"
teacher_id=Session("TId")
finalFilter=Request.Form("finalFilter")

If is_instruct_review Then
	PubTerm=teacher_id&" IN (INSTRUCT_MEMBER1,INSTRUCT_MEMBER2)"
	is_reviewed=toUnsignedInt(Request.Form("In_IS_REVIEWED"))
	If is_reviewed>-1 Then
		PubTerm=PubTerm&" AND (INSTRUCT_MEMBER1="&teacher_id&" AND IS_COMMENT1="&is_reviewed&_
			" OR INSTRUCT_MEMBER2="&teacher_id&" AND IS_COMMENT2="&is_reviewed&")"
	End If
	If Len(finalFilter) Then PubTerm=PubTerm&" AND ("&finalFilter&")"
	table_name="ViewDissertations_instruct"
	arrStatText=Array("未审核","已审核")
Else
	PubTerm="TUTOR_ID="&teacher_id
	If Len(finalFilter) Then PubTerm=PubTerm&" AND ("&finalFilter&")"
	table_name="ViewDissertations_tutor"
End If
If activity_id=-1 Then
	' Dim activity:Set activity=getLastActivityInfoOfStuType(stutypeMBA)
	' If Not activity Is Nothing Then activity_id=activity("Id")
	activity_id = 0
End If
If activity_id>0 Then PubTerm=PubTerm&" AND ActivityId="&activity_id
If teachtype_id>0 Then PubTerm=PubTerm&" AND TEACHTYPE_ID="&teachtype_id
If spec_id>0 Then PubTerm=PubTerm&" AND SPECIALITY_ID="&spec_id
If enter_year>0 Then PubTerm=PubTerm&" AND ENTER_YEAR="&enter_year
If query_task_progress>-1 Then PubTerm=PubTerm&" AND TASK_PROGRESS="&query_task_progress
If query_review_status>-1 Then PubTerm=PubTerm&" AND REVIEW_STATUS="&query_review_status
If is_instruct_review Then
	PubTerm=PubTerm&" ORDER BY SemesterId DESC,ActivityName,GRANT_DEGREE_RESULT,ISINSTRUCTREVIEW"
Else
	PubTerm=PubTerm&" ORDER BY SemesterId DESC,ActivityName,GRANT_DEGREE_RESULT,TEACHTYPE_ID,TASK_PROGRESS,REVIEW_STATUS"
End If
'----------------------PAGE-------------------------
PageNo=""
PageSize=""
If Request.Form("In_PageNo").Count=0 Then
	PageNo=Request.Form("PageNo")
	PageSize=Request.Form("pageSize")
Else
	PageNo=Request.Form("In_PageNo")
	PageSize=Request.Form("In_pageSize")
End If
'------------------------------------------------------
Connect conn
sql="SELECT * FROM "&table_name&" WHERE "&PubTerm
GetRecordSetNoLock conn,rs,sql,count
If IsEmpty(pageSize) Or Not IsNumeric(pageSize) Then
	pageSize=-1
Else
	pageSize=CInt(pageSize)
End If
If pageSize=-1 Then
	If rs.RecordCount>0 Then rs.PageSize=rs.RecordCount
Else
	rs.PageSize=pageSize
End If
pageNo=Request.Form("pageNo")
If IsEmpty(pageNo) Or Not IsNumeric(pageNo) Then
	If rs.PageCount<>0 Then pageNo=1
Else
	pageNo=CInt(pageNo)
	If pageNo>rs.PageCount Then
		If rs.PageCount<>0 Then pageNo=1
	End If
End If
If rs.RecordCount>0 Then rs.AbsolutePage=pageNo
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>查看论文列表</title>
<% useStylesheet "tutor" %>
<% useScript "jquery", "common", "*paper" %>
</head>
<body bgcolor="ghostwhite" onload="return On_Load()">
<center>
<font size=4><b>专业硕士指导论文列表</b></font>
<table width="1000" cellspacing="4" cellpadding="0">
<form id="query_nocheck" method="post" onsubmit="if(Chk_Select())return chkField();else return false">
<tr><td><table cellspacing="4" cellpadding="0">
<tr><td>评阅活动&nbsp;<%=activityList("In_ActivityId", Null, activity_id, True)%></td>
<td><table cellspacing="4" cellpadding="0"><%
Dim ArrayList(2,5),k

FormName="query_nocheck"
k=0
ArrayList(k,0)="学位类别"
ArrayList(k,1)="ViewStudentTypeInfo"
ArrayList(k,2)="TEACHTYPE_ID"
ArrayList(k,3)="TEACHTYPE_NAME"
ArrayList(k,4)=teachtype_id
ArrayList(k,5)="AND TEACHTYPE_ID IN (5,6,7,9)"

k=1
ArrayList(k,0)="专业名称"
ArrayList(k,1)="ViewDissertations"
ArrayList(k,2)="SPECIALITY_ID"
ArrayList(k,3)="SPECIALITY_NAME"
ArrayList(k,4)=spec_id
ArrayList(k,5)=""

k=2
ArrayList(k,0)="年级"
ArrayList(k,1)="ViewStudentInfo"
ArrayList(k,2)="ENTER_YEAR"
ArrayList(k,3)="ENTER_YEAR+'级'"
ArrayList(k,4)=enter_year
ArrayList(k,5)=""
Get_ListJavaMenu ArrayList,k,FormName,""
%></tr></table></td></tr></table></td></tr></table></td></tr><tr><td>
<!--查找-->
<%
	If is_instruct_review Then %>
<select id="is_reviewed" name="In_IS_REVIEWED">
<option value="-1">所有</option>
<option value="0">未审核</option>
<option value="1">已审核</option>
</select><%
	End If %>
<select name="field" onchange="ReloadOperator()">
<option value="s_THESIS_SUBJECT">论文题目</option>
<option value="s_STU_NO">学号</option>
<option value="s_STU_NAME">学生姓名</option>
</select>
<select name="operator">
<script>ReloadOperator()</script>
</select>
<input type="text" name="filter" size="10" onkeypress="checkKey()">
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>">
<input type="submit" value="查找" onclick="genFilter()">
<input type="submit" value="在结果中查找" onclick="genFinalFilter()">
&nbsp;
每页
<select name="pageSize" onchange="if(Chk_Select())submitForm($('#paperList'))">
<option value="-1" <%If pageSize=-1 Then%>selected<%End If%>>全部</option>
<option value="20" <%If rs.PageSize=20 Then%>selected<%End If%>>20</option>
<option value="40" <%If rs.PageSize=40 Then%>selected<%End If%>>40</option>
<option value="60" <%If rs.PageSize=60 Then%>selected<%End If%>>60</option>
</select>
条
&nbsp;
转到
<select name="pageNo" onchange="if(Chk_Select())submitForm($('#paperList'))">
<%
For i=1 to rs.PageCount
	Response.write "<option value="&i
	If rs.AbsolutePage=i Then Response.write " selected"
	Response.write ">"&i&"</option>"
Next
%>
</select>
页
&nbsp;
共<%=rs.RecordCount%>条
</td></tr></form></table>
<form id="paperList" method="post">
<input type="hidden" name="In_ActivityId2" value="<%=activity_id%>">
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>">
<input type="hidden" name="In_SPECIALITY_ID2" value="<%=spec_id%>">
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>">
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>">
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>">
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>">
<input type="hidden" name="pageSize2" value=<%=pageSize%>>
<input type="hidden" name="pageNo2" value=<%=pageNo%>>
<table width="1000" cellpadding="2" cellspacing="1" bgcolor="dimgray">
	<tr bgcolor="gainsboro" align="center" height="25">
		<td width="150" align="center">姓名</td>
		<td width="150" align="center">学号</td>
		<td width="200" align="center">学位类别</td>
		<td align="center">状态</td>
	</tr><%
	Dim is_review_visible,auditor_type,audit_flag,audit_time
	Dim last_rec_activity_id
	For i=1 to rs.PageSize
		If rs.EOF Then Exit For
		is_review_visible=Array(rs("ReviewFileDisplayStatus1")>0,rs("ReviewFileDisplayStatus2")>0)
		substat=vbNullString
		If rs("TASK_PROGRESS")>=tpTbl4Uploaded Then
			stat=rs("STAT_TEXT1")&"，"&rs("STAT_TEXT2")
		ElseIf rs("REVIEW_STATUS")=0 Then
			stat=rs("STAT_TEXT1")
		Else
			stat=rs("STAT_TEXT2")
		End If
		If is_instruct_review Then
			If rs("INSTRUCT_MEMBER1")=teacher_id Then
				auditor_type=1
			Else
				auditor_type=2
			End If
			audit_flag=rs("IsComment"&auditor_type)
			audit_time=rs("AuditTime"&auditor_type)
			If audit_flag Then
				cssclass="paper-status"
				stat=arrStatText(1)
			Else
				audit_time=vbNullString
				cssclass="paper-status-unhandled"
				stat=arrStatText(0)
			End If
		Else
			If rs("REVIEW_STATUS")>=rsReviewed And Not is_review_visible(0) And Not is_review_visible(1) Then
				substat="[评阅结果未开放]"
			End If
			Select Case teacher_id
			Case rs("REVIEWER1")
				auditor_type=0
			Case rs("REVIEWER2")
				auditor_type=1
			Case Else
				auditor_type=-1
			End Select
			reviewer_eval_time=rs("REVIEWER_EVAL_TIME")
			If is_instruct_review Then
				If rs("ISINSTRUCTREVIEW") Then
					cssclass="paper-status-unhandled"
				Else
					cssclass="paper-status"
				End If
			Else
				If rs("ISTABLE") Or rs("ISMODIFY") Or rs("ISEVAL") Or rs("ISINSTRUCTREVIEWDETECT") Or rs("ISDETECT") Then
					cssclass="paper-status-unhandled"
				ElseIf rs("REVIEW_STATUS")=rsAgreedReview And auditor_type<>-1 Then
					If IsNull(reviewer_eval_time) Then
						cssclass="paper-status-unhandled"
					Else
						audit_time=Split(reviewer_eval_time,",")
						If Len(audit_time(auditor_type))=0 Then
							cssclass="paper-status-unhandled"
						Else
							cssclass="paper-status"
						End If
					End If
				Else
					cssclass="paper-status"
				End If
			End If
		End If
		stu_type_name=rs("TEACHTYPE_NAME")
		If rs("TEACHTYPE_ID")=5 Then
			stu_type_name=Format("{0}（{1}）",stu_type_name,rs("SPECIALITY_NAME"))
		End If
		is_granted_degree=rs("GRANT_DEGREE_RESULT")
		If last_rec_activity_id<>rs("ActivityId") Then
			' 评阅活动分组
			last_rec_activity_id=rs("ActivityId")
	%><tr>
		<td class="paper-group" colspan="5"><%=rs("ActivityName")%></td>
	</tr><%
		End If
	%><tr bgcolor="ghostwhite" height="30">
		<td align="center"><a href="#" onclick="return showPaperDetail(<%=rs("ID")%>,2)"><%=rs("STU_NAME")%></a></td>
		<td align="center"><%=rs("STU_NO")%></td>
		<td align="center"><%=stu_type_name%></td>
		<td align="center"><%
		If is_granted_degree Then %>
		<span class="paper-substate">已毕业，完成评阅系统流程</span><%
		Else %>
		<a href="#" onclick="return showPaperDetail(<%=rs("ID")%>,2)"><span class="<%=cssclass%>"><%=stat%></span></a><%
		End If
		If Len(substat) Then
		%><br/><span class="review-display-status"><%=substat%></span><%
		End If %></a></td></tr><%
		rs.MoveNext()
	Next
%></table></form></center>
<script type="text/javascript">
	$("#is_reviewed").val("<%=is_reviewed%>");
</script></body></html><%
	CloseRs rs
	CloseConn conn
%>
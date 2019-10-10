<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Tid")) Then Response.Redirect("../error.asp?timeout")

Dim bModified,PubTerm,PageNo,PageSize
bModified=Request.QueryString("modified")="1"
activity_id=toUnsignedInt(Request.Form("In_ActivityId"))
teachtype_id=toUnsignedInt(Request.Form("In_TEACHTYPE_ID"))
spec_id=toUnsignedInt(Request.Form("In_SPECIALITY_ID"))
enter_year=toUnsignedInt(Request.Form("In_ENTER_YEAR"))
query_task_progress=toUnsignedInt(Request.Form("In_TASK_PROGRESS"))
query_review_status=toUnsignedInt(Request.Form("In_REVIEW_STATUS"))
Tid=Session("Tid")
finalFilter=Request.Form("finalFilter")
If Len(finalFilter) Then PubTerm="AND ("&finalFilter&")"
If activity_id=-1 Then
	Dim activity:Set activity=getLastActivityInfoOfStuType(stutypeMBA)
	If Not IsNull(activity) Then activity_id=activity("Id")
End If
If activity_id>0 Then PubTerm=PubTerm&" AND ActivityId="&activity_id
If teachtype_id>0 Then PubTerm=PubTerm&" AND TEACHTYPE_ID="&teachtype_id
If spec_id>0 Then PubTerm=PubTerm&" AND SPECIALITY_ID="&spec_id
If enter_year>0 Then PubTerm=PubTerm&" AND ENTER_YEAR="&enter_year
If query_task_progress>-1 Then PubTerm=PubTerm&" AND TASK_PROGRESS="&query_task_progress
If query_review_status>-1 Then PubTerm=PubTerm&" AND REVIEW_STATUS="&query_review_status

If bModified Then
	PubTerm=PubTerm&" AND REVIEW_STATUS>="&rsModifyThesisUploaded
	table_title="修改后专业硕士指导论文列表"
Else
	table_title="专业硕士指导论文列表"
End If
PubTerm=PubTerm&" ORDER BY ISTABLE DESC,ISMODIFY DESC,ISEVAL DESC,ISREVIEW DESC,ISDETECT DESC"
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
sql="SELECT * FROM ViewDissertations_tutor WHERE TUTOR_ID="&Tid&PubTerm
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
<% useScript "jquery", "common", "thesis" %>
</head>
<body bgcolor="ghostwhite" onload="return On_Load()">
<center>
<font size=4><b><%=table_title%></b></font>
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
ArrayList(k,1)="ViewThesisInfo"
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
%></tr></table></td></tr></table></td></tr>
<tr><td><table cellspacing="4" cellpadding="0"><tr><td>表格审核状态</td><td><select name="In_TASK_PROGRESS"><option value="-1">请选择</option><%
GetMenuListPubTerm "ReviewStatuses","STATUS_ID1","STATUS_NAME",query_task_progress,"AND STATUS_ID1 IS NOT NULL"
%></select></td><td>论文审核状态</td><td><select name="In_REVIEW_STATUS"><option value="-1">请选择</option><%
GetMenuListPubTerm "ReviewStatuses","STATUS_ID2","STATUS_NAME",query_review_status,"AND STATUS_ID2 IS NOT NULL"
%></select></td></tr></table></td></tr><tr><td>
<!--查找-->
<select name="field" onchange="ReloadOperator()">
<option value="s_THESIS_SUBJECT">论文题目</option>
<option value="s_STU_NO">学号</option>
<option value="s_STU_NAME">学生姓名</option>
<option value="ms_REVIEW_RESULT_TEXT1_tutor|REVIEW_RESULT_TEXT2_tutor">送审结果</option>
<option value="s_FINAL_RESULT_TEXT_tutor">处理意见</option>
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
<select name="pageSize" onchange="if(Chk_Select())submitForm($('#fmThesisList'))">
<option value="-1" <%If pageSize=-1 Then%>selected<%End If%>>全部</option>
<option value="20" <%If rs.PageSize=20 Then%>selected<%End If%>>20</option>
<option value="40" <%If rs.PageSize=40 Then%>selected<%End If%>>40</option>
<option value="60" <%If rs.PageSize=60 Then%>selected<%End If%>>60</option>
</select>
条
&nbsp;
转到
<select name="pageNo" onchange="if(Chk_Select())submitForm($('#fmThesisList'))">
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
<form id="fmThesisList" method="post">
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
		<td align="center">论文题目</td>
		<td width="80" align="center">姓名</td>
		<td width="90" align="center">学号</td>
		<td width="120" align="center">专业</td>
		<td width="50" align="center">学位类别</td><%
		If Not bModified Then %>
		<td width="80" align="center">送审结果1</td>
		<td width="80" align="center">送审结果2</td><%
		End If %>
		<td width="80" align="center">处理意见</td>
		<td width="180" align="center">状态</td>
	</tr><%
	Dim bIsReviewVisible
	For i=1 to rs.PageSize
		If rs.EOF Then Exit For
		bIsReviewVisible=rs("REVIEW_FILE_STATUS")=1 Or rs("REVIEW_FILE_STATUS")=3
		substat=vbNullString
		If rs("TASK_PROGRESS")>=tpTbl4Uploaded Then
			stat=rs("STAT_TEXT1")&"，"&rs("STAT_TEXT2")
		ElseIf rs("REVIEW_STATUS")=0 Then
			stat=rs("STAT_TEXT1")
		Else
			stat=rs("STAT_TEXT2")
		End If
		If Not bIsReviewVisible And rs("REVIEW_STATUS")>=rsReviewed Then
			substat="[评阅结果未开放]"
		End If
		Select Case Tid
		Case rs("REVIEWER1")
			reviewer=0
		Case rs("REVIEWER2")
			reviewer=1
		Case Else
			reviewer=-1
		End Select
		reviewer_eval_time=rs("REVIEWER_EVAL_TIME")
		If rs("ISTABLE") Or rs("ISMODIFY") Or rs("ISEVAL") Or rs("ISREVIEW") Or rs("ISDETECT") Then
			cssclass="thesisstat_unhandled"
		ElseIf rs("REVIEW_STATUS")=rsAgreeReview And reviewer<>-1 Then
			If IsNull(reviewer_eval_time) Then
				cssclass="thesisstat_unhandled"
			Else
			review_time=Split(reviewer_eval_time,",")
			If Len(review_time(reviewer))=0 Then
				cssclass="thesisstat_unhandled"
			Else
				cssclass="thesisstat"
			End If
		End If
	Else
			cssclass="thesisstat"
		End If
	%><tr bgcolor="ghostwhite" height="30">
		<td align="center"><a href="#" onclick="return showThesisDetail(<%=rs("ID")%>,2)"><%=HtmlEncode(rs("THESIS_SUBJECT"))%></a></td>
		<td align="center"><%=HtmlEncode(rs("STU_NAME"))%></td>
		<td align="center"><%=rs("STU_NO")%></td>
		<td align="center"><%=HtmlEncode(rs("SPECIALITY_NAME"))%></td>
		<td align="center"><%=rs("TEACHTYPE_NAME")%></td><%
		If Not bModified Then %>
		<td align="center"><%=rs("REVIEW_RESULT_TEXT1_tutor")%></td>
		<td align="center"><%=rs("REVIEW_RESULT_TEXT2_tutor")%></td><%
	End If %>
		<td align="center"><%=rs("FINAL_RESULT_TEXT_tutor")%>
		<td align="center"><a href="#" onclick="return showThesisDetail(<%=rs("ID")%>,2)"><span class="<%=cssclass%>"><%=stat%></span></a><%
		If Len(substat) Then
		%><br/><span class="thesissubstat"><%=substat%></span><%
		End If %></a></td></tr><%
		rs.MoveNext()
	Next
%></table></form></center></body></html><%
	CloseRs rs
	CloseConn conn
%>
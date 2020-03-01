<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")

Dim PubTerm,PageNo,PageSize
teacher_id=Session("TId")
activity_id=toUnsignedInt(Request.Form("In_ActivityId"))
teachtype_id=toUnsignedInt(Request.Form("In_TEACHTYPE_ID"))
spec_id=toUnsignedInt(Request.Form("In_SPECIALITY_ID"))
is_reviewed=toUnsignedInt(Request.Form("In_IS_REVIEWED"))
finalFilter=Request.Form("finalFilter")
If Len(finalFilter) Then PubTerm="AND ("&finalFilter&")"
If activity_id=-1 Then
	Dim semester_info:semester_info=getCurrentSemester()
	PubTerm=PubTerm&" AND SemesterId="&semester_info(3)
	activity_id=0
End If
If teachtype_id>0 Then PubTerm=PubTerm&" AND TEACHTYPE_ID="&teachtype_id
If spec_id>0 Then PubTerm=PubTerm&" AND SPECIALITY_ID="&spec_id
If is_reviewed>-1 Then
	PubTerm=PubTerm&" AND (REVIEWER1="&teacher_id&" AND IS_REVIEWER_EVAL1="&is_reviewed&_
		" OR REVIEWER2="&teacher_id&" AND IS_REVIEWER_EVAL2="&is_reviewed&")"
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
arrStatText=Array("未评阅","已评阅")
Connect conn
sql="SELECT * FROM ViewDissertations_expert WHERE "&teacher_id&" IN (REVIEWER1,REVIEWER2) "&PubTerm
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
If rs.RecordCount>0 Then
	rs.AbsolutePage=pageNo
End If
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>查看论文列表</title>
<% useStylesheet "tutor" %>
<% useScript "jquery", "common", "paper" %>
</head>
<body class="exp" onload="return On_Load()">
<center><div class="content">
<font size=4><b>专业硕士评阅论文列表</b></font><%
If Not checkIfProfileFilledIn() Then
%><p><span class="tip">您尚未完善个人信息，<a href="profile.asp">请点击这里编辑。</a></span></p><%
End If %>
<table width="1000" cellspacing="4" cellpadding="0">
<form id="query_nocheck" method="post" onsubmit="if(Chk_Select())return chkField();else return false">
<tr><td><table cellspacing="4" cellpadding="0">
<tr><td>评阅活动&nbsp;<%=activityList("In_ActivityId", Null, activity_id, True)%></td>
<td><table cellspacing="4" cellpadding="0"><%
Dim ArrayList(1,5),k

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
Get_ListJavaMenu ArrayList,k,FormName,""
%></tr></table></td>
<td>评阅状态</td>
<td><select id="is_reviewed" name="In_IS_REVIEWED">
<option value="-1">所有</option>
<option value="0">未评阅</option>
<option value="1">已评阅</option>
</select></td></tr></table></td></tr><tr><td>
<!--查找-->
<select name="field" onchange="ReloadOperator()">
<option value="s_THESIS_SUBJECT">论文题目</option>
<option value="s_SPECIALITY_NAME">专业</option>
<option value="s_TEACHTYPE_NAME">学位类别</option>
</select>
<select name="operator">
<script>ReloadOperator()</script>
</select>
<input type="text" name="filter" size="10" onkeypress="checkKey()">
<input type="hidden" name="finalFilter" value="<%=finalFilter%>">
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
	Response.Write "<option value="&i
	If rs.AbsolutePage=i Then Response.Write " selected"
	Response.Write ">"&i&"</option>"
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
<input type="hidden" name="finalFilter2" value="<%=finalFilter%>">
<input type="hidden" name="pageNo2" value=<%=PageNo%>>
<input type="hidden" name="pageSize2" value=<%=PageSize%>>
<table width="1000" cellpadding="2" cellspacing="1" bgcolor="dimgray">
	<tr bgcolor="gainsboro" align="center" height="25">
		<td align="center">论文题目</td>
		<td width="120" align="center">专业</td>
		<td width="50" align="center">学位类别</td>
		<td align="center">状态</td>
		<td width="150" align="center">评阅时间</td>
	</tr><%
	Dim arr,review_flag,review_time,review_result
	For i=1 to rs.PageSize
		If rs.EOF Then Exit For
		If rs("REVIEWER1")=teacher_id Then
			reviewer_type=0
		Else
			reviewer_type=1
		End If
		If Not IsNull(rs("REVIEW_RESULT")) Then
			review_result=Split(rs("REVIEW_RESULT"),",")
		Else
			ReDim review_result(2)
		End If
		review_flag=rs("IS_REVIEWER_EVAL"&(reviewer_type+1))
		If review_flag Then
			arr=Split(rs("REVIEWER_EVAL_TIME"),",")
			review_time=toDateTime(arr(reviewer_type),1)&" "&toDateTime(arr(reviewer_type),4)
			cssclass="paper-status"
			stat=arrStatText(1)
		Else
			review_time=vbNullString
			cssclass="paper-status-unhandled"
			stat=arrStatText(0)
		End If
%><tr bgcolor="ghostwhite" height="30">
		<td align="center"><a href="#" onclick="return showPaperDetail(<%=rs("ID")%>,3)"><%=HtmlEncode(rs("THESIS_SUBJECT"))%></a></td>
		<td align="center"><%=HtmlEncode(rs("SPECIALITY_NAME"))%></td>
		<td align="center"><%=rs("TEACHTYPE_NAME")%></td>
		<td align="center"><a href="#" onclick="return showPaperDetail(<%=rs("ID")%>,3)"><span class="<%=cssclass%>"><%=stat%></span></a></td>
		<td align="center"><%=review_time%></td></tr><%
		rs.MoveNext()
	Next
%></table></form></div></center>
<script type="text/javascript">
	$("#is_reviewed").val("<%=is_reviewed%>");
</script>
</body></html><%
	CloseRs rs
	CloseConn conn
%>
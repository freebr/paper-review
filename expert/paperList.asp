<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")

Dim PubTerm,PageNo,PageSize
teacher_id=Session("TId")
activity_id=toUnsignedInt(Request.Form("In_ActivityId"))
teachtype_id=toUnsignedInt(Request.Form("In_TEACHTYPE_ID"))
is_reviewed=toUnsignedInt(Request.Form("In_IS_REVIEWED"))
finalFilter=Request.Form("finalFilter")
If Len(finalFilter) Then PubTerm="AND ("&finalFilter&")"

ConnectDb conn
If IsEmpty(activity_id) Or activity_id=-1 Then
	' 获取专家待评阅的学生类型
	sql="SELECT dbo.getUnhandledReviewType(?)"
	Set ret=ExecQuery(conn,sql,CmdParam("expert_id",adInteger,4,teacher_id))
	unhandled_review_type=ret("rs")(0)
	CloseRs rs
	Dim activity:Set activity=getLastActivityInfoOfStuType(unhandled_review_type)
	If activity Is Nothing Then
		CloseConn conn
		showErrorPage "您暂时没有需要评阅的论文，请稍后再查看！", "提示"
	End If
	activity_id=activity("Id")
	PubTerm=PubTerm&" AND SemesterId="&activity("SemesterId")
End If
If teachtype_id>0 Then PubTerm=PubTerm&" AND TEACHTYPE_ID="&teachtype_id
If is_reviewed>-1 Then
	PubTerm=PubTerm&" AND (REVIEWER1="&teacher_id&" AND IsComment1="&is_reviewed&_
		" OR REVIEWER2="&teacher_id&" AND IsComment2="&is_reviewed&")"
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
arrStatText=Array("待评阅","已评阅")
sql="SELECT dbo.getUnhandledReviewPaperCount(?,?)"
Set ret=ExecQuery(conn,sql,CmdParam("expert_id",adInteger,4,teacher_id),_
	CmdParam("activity_id",adInteger,4,activity_id))
count_unhandled=ret("rs")(0)
CloseRs rs

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
<font size=4><b><%=activity("Name")%>专业硕士评阅论文列表</b></font><%
If Not checkIfProfileFilledIn() Then
%><p><span class="tip">您尚未完善个人信息，<a href="profile.asp">请点击这里编辑。</a></span></p><%
End If %>
<p align="center"><span class="tip">您本次共需评阅<%=rs.RecordCount%>篇论文，其中<%=count_unhandled%>篇待评阅。</span></p>
<table width="1000" cellspacing="4" cellpadding="0" style="display: none">
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
Get_ListJavaMenu ArrayList,k,FormName,""
%></tr></table></td>
<!--<td>评阅状态</td>
<td><select id="is_reviewed" name="In_IS_REVIEWED">
<option value="-1">所有</option>
<option value="0">未评阅</option>
<option value="1">已评阅</option>
</select></td>-->
</tr></table></td></tr><tr><td>
<!--查找-->
<select name="field" onchange="ReloadOperator()">
<option value="s_THESIS_SUBJECT">论文题目</option>
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
<input type="hidden" name="finalFilter2" value="<%=finalFilter%>">
<input type="hidden" name="pageNo2" value=<%=PageNo%>>
<input type="hidden" name="pageSize2" value=<%=PageSize%>>
<table width="1000" cellpadding="2" cellspacing="1" bgcolor="dimgray">
	<tr bgcolor="gainsboro" align="center" height="25">
		<td width="500" align="center">论文题目</td>
		<td width="200" align="center">学位类别</td>
		<td align="center">状态</td>
		<td width="150" align="center">评阅时间</td>
	</tr><%
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
	is_reviewed=rs("IsComment"&(reviewer_type+1))
	If is_reviewed Then
		arr=Split(rs("REVIEWER_EVAL_TIME"),",")
		review_time=toDateTime(arr(reviewer_type),1)&" "&toDateTime(arr(reviewer_type),4)
		cssclass="paper-status"
		stat=arrStatText(1)
		show_link=False
	Else
		review_time=vbNullString
		cssclass="paper-status-unhandled"
		stat=arrStatText(0)
		show_link=True
	End If
	stu_type_name=rs("TEACHTYPE_NAME")
	If rs("TEACHTYPE_ID")=5 Then
		stu_type_name=Format("{0}（{1}）",stu_type_name,rs("SPECIALITY_NAME"))
	End If
%><tr bgcolor="ghostwhite" height="30">
		<td align="center"><a href="#" onclick="return showPaperDetail(<%=rs("ID")%>,3)"><%=HtmlEncode(rs("THESIS_SUBJECT"))%></a></td>
		<td align="center"><%=stu_type_name%></td>
		<td align="center"><%
	If show_link Then %>
		<a href="#" onclick="return showPaperDetail(<%=rs("ID")%>,3)"><span class="<%=cssclass%>"><%=stat%></span></a><%
	Else %>
		<span class="<%=cssclass%>"><%=stat%></span><%
	End If %>
		</td>
		<td align="center"><%=review_time%></td></tr><%
		rs.MoveNext()
Next
%></table></form></div></center>
<script type="text/javascript">
	$("#is_reviewed").val("<%=is_reviewed%>");
</script></body></html><%
CloseRs rs
CloseConn conn
%>
<%Response.Charset="utf-8"
Response.Expires=-1%>
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
teachertype=Request.QueryString("type")
If IsEmpty(teachertype) Then teachertype=2
ctrlname1=Request.QueryString("ctrl1")
ctrlname2=Request.QueryString("ctrl2")
itemid=Request.QueryString("item")

Dim PubTerm,PageNo,PageSize
finalFilter=Request.Form("finalFilter")
If Len(finalFilter) Then PubTerm="AND ("&finalFilter&")"

'----------------------PAGE-------------------------
PageNo=""
PageSize=""
If Request.Form("In_PageNo").Count=0 Then
	PageNo=Request.Form("pageNo")
	PageSize=Request.Form("pageSize")
Else
	PageNo=Request.Form("In_pageNo")
	PageSize=Request.Form("In_pageSize")
End If
bShowAll=Request.QueryString="showAll"
If bShowAll Then PageSize=-1
'------------------------------------------------------

Connect conn
If teachertype=1 Then	' 校内导师
	sql="SELECT * FROM VIEW_TUTOR_LIST_GROUPBY_TEACHER A LEFT JOIN VIEW_TEACHER_INFO B ON A.TEACHER_ID=B.TEACHERID WHERE 1=1 "&PubTerm&" ORDER BY TEACHER_NAME"
	title="校内导师名单"
Else	' 校外专家
	sql="SELECT * FROM VIEW_TEST_THESIS_REVIEW_EXPERT_INFO WHERE INSCHOOL=0 AND VALID=1 "&PubTerm&" ORDER BY EXPERT_NAME"
	title="校外专家名单"
End If
GetRecordSetNoLock conn,rs,sql,result
If IsEmpty(pageSize) Or Not IsNumeric(pageSize) Then
  pageSize=60
Else
	pageSize=CInt(pageSize)
End If
If pageSize=-1 Then
	If rs.RecordCount>0 Then rs.PageSize=rs.RecordCount
Else
  rs.PageSize=pageSize
End If
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
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.11.3.min.js" type="text/javascript"></script>
<script src="../scripts/query.js" type="text/javascript"></script>
<script src="../scripts/expertList.js" type="text/javascript"></script>
<meta name="theme-color" content="#2D79B2" />
<title>选择专家</title>
</head>
<body bgcolor="ghostwhite" onload="return On_Load()">
<center>
<font size=4><b><%=title%></b></font>
<p align="center"><input type="button" id="btnlist1" value="校内导师" />&emsp;<input type="button" id="btnlist2" value="校外专家" /></p>
<table width="800" cellpadding="2" cellspacing="1" bgcolor="dimgray">
	<form id="query" method="post" onsubmit="return chkField()">
	<tr bgcolor="ghostwhite"><td colspan=7>
	<!--查找-->
	<select name="field" onchange="ReloadOperator()">
	<option value="s_EXPERT_NAME">专家姓名</option>
	<option value="s_PRO_DUTY_NAME">职称</option>
	<option value="s_EXPERTISE">学科专长</option>
	<option value="s_WORKPLACE">单位（住址）</option>
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
	<select name="pageSize" onchange="submitForm(this.form)">
	<option value="-1" <%If pageSize=-1 Then%>selected<%End If%>>全部</option>
	<option value="20" <%If rs.PageSize=20 Then%>selected<%End If%>>20</option>
	<option value="40" <%If rs.PageSize=40 Then%>selected<%End If%>>40</option>
	<option value="60" <%If rs.PageSize=60 Then%>selected<%End If%>>60</option>
	</select>
	条
	&nbsp;
	转到
	<select name="pageNo" onchange="submitForm(this.form)">
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
	<input type="button" value="显示全部" onclick="showAllRecords(this.form)"></td></tr>
  <tr bgcolor="gainsboro" align="center" height="25"><%
If teachertype=1 Then
%><td width="100" align=center>姓名</td>
  <td width="80" align=center>职称</td>
  <td align=center>所在院系</td>
	<td width="140" align=center>单位（住址）</td>
	<td width="100" align=center>联系电话</td>
	<td width="120" align=center>邮箱</td>
	<td width="80" align=center>选择</td><%
Else
%><td width="100" align=center>姓名</td>
  <td width="80" align=center>职称</td>
  <td align=center>学科专长</td>
	<td width="140" align=center>单位（住址）</td>
	<td width="100" align=center>联系电话</td>
	<td width="120" align=center>邮箱</td>
	<td width="80" align=center>选择</td><%
End If
%></tr><%
If teachertype=1 Then
  For i=1 To pageSize
  	If rs.EOF Then Exit For
%>
<tr bgcolor="ghostwhite">
  <td align=center><%=HtmlEncode(rs("TEACHER_NAME"))%></td>
  <td align=center><%=HtmlEncode(rs("PRO_DUTYNAME"))%></td>
  <td align=center><%=HtmlEncode(rs("DEPT_NAME"))%></td>
  <td align=center><%=HtmlEncode(rs("OFFICE_ADDRESS"))%></td>
  <td align=center><%=HtmlEncode(rs("MOBILE"))%></td>
  <td align=center><%=HtmlEncode(rs("EMAIL"))%></td>
  <td align=center><a href="#" onclick="selectItem('<%=toJsString(rs("TEACHER_NAME"))%>',<%=rs("TEACHER_ID")%>)">选择</a></td>
</td></tr><%
  	rs.MoveNext()
  Next
Else
  For i=1 To pageSize
  	If rs.EOF Then Exit For%>
<tr bgcolor="ghostwhite">
  <td align=center><%=HtmlEncode(rs("EXPERT_NAME"))%></td>
  <td align=center><%=HtmlEncode(rs("PRO_DUTY_NAME"))%></td>
  <td align=center><%=HtmlEncode(rs("EXPERTISE"))%></td>
  <td align=center><%=HtmlEncode(rs("WORKPLACE"))%></td>
  <td align=center><%=HtmlEncode(rs("MOBILE"))%></td>
  <td align=center><%=HtmlEncode(rs("EMAIL"))%></td>
  <td align=center><a href="#" onclick="selectItem('<%=toJsString(rs("EXPERT_NAME"))%>',<%=rs("TEACHER_ID")%>)">选择</a></td>
</td></tr><%
  	rs.MoveNext()
  Next
End If
%></form></table></center></body>
<script type="text/javascript"><%
	For i=1 To 2 %>
	$('#btnlist<%=i%>').click(function() {
		location.href="?type=<%=i%>&ctrl1=<%=ctrlname1%>&ctrl2=<%=ctrlname2%>&item=<%=itemid%>";
	});<%
	Next %>
	function selectItem(val1,val2) {
		opener.$('[name="<%=ctrlname1%>"]').eq(<%=itemid%>).val(val1);
		opener.$('[name="<%=ctrlname2%>"]').eq(<%=itemid%>).val(val2);
		window.close();
		return;
	}
</script></html><%
CloseRs rs
CloseConn conn %>
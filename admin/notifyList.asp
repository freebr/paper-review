<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")

Dim PubTerm,PageNo,PageSize
sem_info=getCurrentSemester()
user_type=Request.Form("usertype")
finalFilter=Request.Form("finalFilter")
If Len(finalFilter) Then PubTerm="AND ("&finalFilter&")"
If Len(user_type) And user_type<>"0" Then
	PubTerm=PubTerm&" AND USER_TYPE="&toSqlString(user_type)
End If
PubTerm=PubTerm&" AND AUDIT_COUNT+REQUEST_REVIEW_COUNT+INFO_IMPORTED_COUNT>0"
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
sql="SELECT * FROM ViewNotifyInfo WHERE 1=1 "&PubTerm&" ORDER BY USER_TYPE,USER_NAME"
GetRecordSetNoLock conn,rs,sql,count
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
<% useStylesheet "admin" %>
<% useScript "jquery", "common", "notifyList" %>
</head>
<body>
<center>
<font size=4><b>邮件/短信通知列表</b></font>
<table cellspacing=4 cellpadding=0>
<form id="query" method="post" onsubmit="return chkField()">
<tr><td colspan=2>
<!--查找-->
<select name="field" onchange="ReloadOperator()">
<option value="s_USER_NAME">姓名</option>
<option value="d_LAST_NOTIFY_TIME">上次通知时间</option>
<option value="d_LAST_ACTIVE_TIME">最后使用系统时间</option>
</select>
<select name="operator">
<script>ReloadOperator()</script>
</select>
<input type="text" name="filter" size="10" onkeypress="checkKey()">
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>">
<input type="submit" value="查找" onclick="genFilter()">
<input type="submit" value="在结果中查找" onclick="genFinalFilter()"><%
If Len(PubTerm) Then %>
&nbsp;
每页
<select name="pageSize" onchange="submitForm($(this.form))">
<option value="-1" <%If pageSize=-1 Then%>selected<%End If%>>全部</option>
<option value="20" <%If rs.PageSize=20 Then%>selected<%End If%>>20</option>
<option value="40" <%If rs.PageSize=40 Then%>selected<%End If%>>40</option>
<option value="60" <%If rs.PageSize=60 Then%>selected<%End If%>>60</option>
</select>
条
&nbsp;
转到
<select name="pageNo" onchange="submitForm($(this.form))">
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
共<%=rs.RecordCount%>条<%
End If %>
<input type="button" value="显示全部" onclick="showAllRecords($(this.form))">
&nbsp;全选<input type="checkbox" onclick="checkAll()" id="chk" /></td></tr>
<tr><td colspan=2><input type="button" value="更新收件人列表" onclick="submitForm($('#fmNotifyList'),'refreshNotifyList.asp')" />
<input type="button" value="显示待通知的导师列表" onclick="showNotifyList($('#fmNotifyList'),1)" />
<input type="button" value="显示待通知的专家列表" onclick="showNotifyList($('#fmNotifyList'),2)" />
<input type="button" value="批量发送通知" onclick="notifySelection($('#fmNotifyList'))" /></form>
<form id="fmNotifyList" method="post">
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>">
<input type="hidden" name="pageSize2" value=<%=pageSize%>>
<input type="hidden" name="pageNo2" value=<%=pageNo%>>
<input type="hidden" name="usertype" id="usertype" value="0">
<table width="1000" cellpadding="2" cellspacing="1" bgcolor="dimgray">
  <tr bgcolor="gainsboro" height="25">
    <td width="100" align="center">通知对象</td>
    <td width="50" align="center">身份</td>
    <td align="center">待发送通知邮件情况</td>
    <td width="100" align="center">上次通知日期</td>
    <td width="100" align="center">上次通知结果</td>
    <td width="120" align="center">最后使用评阅系统时间</td>
    <td width="30" align="center">选择</td>
    <td width="100" align="center">操作</td>
  </tr><%
  Dim arrUserType:arrUserType=Array("","导师","评阅专家","学生")
  Dim arrNotifyResult:arrNotifyResult=Array("失败","成功")
  Dim user_type,notify_desc,last_notify_result
  For i=1 to rs.PageSize
    If rs.EOF Then Exit For
    user_type=arrUserType(rs("USER_TYPE"))
    notify_desc=vbNullString
    last_notify_result=vbNullString
    If rs("AUDIT_COUNT")>0 Then
    	Select Case rs("USER_TYPE")
    	Case 1
    		notify_desc="待审核："&rs("AUDIT_COUNT")
     	Case 2
    		notify_desc="待通知评阅："&rs("AUDIT_COUNT")
  		End Select
 		End If
    If rs("REQUEST_REVIEW_COUNT")>0 Then
    	If Len(notify_desc) Then notify_desc=notify_desc&"，"
    	notify_desc=notify_desc&"待通知送审："&rs("REQUEST_REVIEW_COUNT")
  	End If
    If rs("INFO_IMPORTED_COUNT")>0 Then
    	If Len(notify_desc) Then notify_desc=notify_desc&"，"
    	notify_desc=notify_desc&"待通知答辩安排/意见："&rs("INFO_IMPORTED_COUNT")
  	End If
  	If Not IsNull(rs("LAST_NOTIFY_MAIL_RESULT")) Then
	  	Select Case rs("USER_TYPE")
	  	Case 1,3
	  		last_notify_result=arrNotifyResult(Abs(rs("LAST_NOTIFY_MAIL_RESULT")))
	  	Case 2
	  		last_notify_result="邮件："&arrNotifyResult(Abs(rs("LAST_NOTIFY_MAIL_RESULT")))&_
	  											 "<br/>手机："&arrNotifyResult(Abs(rs("LAST_NOTIFY_MOB_RESULT")))
	  	End Select
	  End If
  %><tr bgcolor="ghostwhite">
    <td align="center"><a href="#" onclick="return showTeacherResume('<%=rs("USER_ID")%>')"><%=HtmlEncode(rs("USER_NAME"))%></a></td>
    <td align="center"><%=user_type%></td>
    <td align="center"><%=notify_desc%></td>
    <td align="center"><%=HtmlEncode(rs("LAST_NOTIFY_TIME"))%></td>
    <td align="center"><%=last_notify_result%></td>
    <td align="center"><%=HtmlEncode(rs("LAST_ACTIVE_TIME"))%></td>
    <td align="center"><input type="checkbox" name="sel" value="<%=rs("USER_TYPE")%>.<%=rs("USER_ID")%>" /></td>
    <td align="center"><input type="button" value="发送通知" onclick="notify($(this.form),<%=rs("USER_TYPE")%>,<%=rs("USER_ID")%>)" /></td></tr><%
  	rs.MoveNext()
  Next
%></table></form></center></body></html><%
  CloseRs rs
  CloseConn conn
%>
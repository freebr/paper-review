<%Response.Charset="utf-8"
Response.Expires=-1%>
<!--#include file="../inc/db.asp"-->
<!--#include virtual="/pub/mail.asp"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("user")) Then Response.Redirect("../error.asp?timeout")
ids=Request.Form("sel")
If Len(ids)=0 Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">请选择论文记录！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End
End If

teachtype_id=Request.Form("In_TEACHTYPE_ID2")
class_id=Request.Form("In_CLASS_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
Connect conn
sql="SELECT * FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE ID IN ("&ids&")"
GetRecordSetNoLock conn,rs,sql,result
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
<script src="../scripts/query.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite">
<center>
<font size=4><b>为<%=result%>篇送审论文匹配评阅专家（单击方框选择）</b></font>
<form id="fmChooseExp" method="post" action="chooseExpert.asp?step=2">
<input type="hidden" name="ids" value="<%=ids%>" />
<input type="hidden" name="In_TEACHTYPE_ID" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize" value="<%=pageSize%>" />
<input type="hidden" name="pageNo" value="<%=pageNo%>" />
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID2" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" />
<table width="800" cellpadding="2" cellspacing="1" bgcolor="dimgray">
  <tr bgcolor="gainsboro" height="25">
    <td align=center>论文题目</td>
    <td width="80" align=center>姓名</td>
    <td width="90" align=center>学号</td>
    <td width="120" align=center>专业</td>
    <td width="50" align=center>学位类别</td>
    <td width="60" align=center>导师</td>
		<td width="110" align=center>状态</td>
  </tr><%
  Dim review_result
  For i=1 to rs.PageSize
    If rs.EOF Then Exit For
    substat=vbNullString
		If rs("TASK_PROGRESS")>=tpTbl4Uploaded Then
    	stat=rs("TASK_PROGRESS_NAME")&"，"&rs("REVIEW_STATUS_NAME")
		ElseIf rs("REVIEW_STATUS")=0 Then
    	stat=rs("TASK_PROGRESS_NAME")
    Else
    	stat=rs("REVIEW_STATUS_NAME")
    	If rs("REVIEW_STATUS")>=rsReviewed And rs("REVIEW_FILE_STATUS")<>3 Then
    		substat="评阅结果["&arrReviewFileStat(rs("REVIEW_FILE_STATUS"))&"]"
    	End If
  	End If
  %><tr bgcolor="ghostwhite">
    <td align=center><%=HtmlEncode(rs("THESIS_SUBJECT"))%></td>
    <td align=center><%=HtmlEncode(rs("STU_NAME"))%></td>
    <td align=center><%=rs("STU_NO")%></td>
    <td align=center><%=HtmlEncode(rs("SPECIALITY_NAME"))%></td>
    <td align=center><%=rs("TEACHTYPE_NAME")%></td>
    <td align=center><%=HtmlEncode(rs("TUTOR_NAME"))%></td>
    <td align=center><span class="thesisstat"><%=stat%></span><%
    If Len(substat) Then
    %><br/><span class="thesissubstat"><%=substat%></span><%
    End If %></td></tr><%
  	rs.MoveNext()
  Next
%></table>
<p><font size=4><b>请选择要匹配的评阅专家</b></font></p>
<table class="tblform" width="800" cellpadding="2" cellspacing="1" bgcolor="dimgray">
<tr bgcolor="gainsboro" align="center" height="25">
<td width="100" align=center>专家一：</td>
<td width="200" align=center><input type="text" class="selectbox" name="expertname" size=20 value="单击选择..." onclick="window.open('selectExpert.asp?ctrl1=expertname&ctrl2=expertid&item=0','','width=800,height=500,location=no,scrollbars=yes')"/><input type="hidden" name="expertid" /></td>
<td width="100" align=center>专家二：</td>
<td width="200" align=center><input type="text" class="selectbox" name="expertname" size=20 value="单击选择..." onclick="window.open('selectExpert.asp?ctrl1=expertname&ctrl2=expertid&item=1','','width=800,height=500,location=no,scrollbars=yes')"/><input type="hidden" name="expertid" /></td>
</tr></table><p><input type="submit" name="btnsubmit" value="确 定" />&emsp;
<input type="submit" name="btnreturn" value="返 回" onclick="this.form.action='thesisList.asp'" /></p></form></center></body>
<script type="text/javascript">
	$('#btnsubmit').click(function(){
		$(this).val('正在提交，请稍候……').attr('disabled',true);
		this.form.submit();
	}).attr('disabled',false);
</script></html><%
  CloseRs rs
  CloseConn conn
%>
<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<!--#include file="../inc/setEditor.asp"-->
<!--#include file="../inc/ckeditor/ckeditor.asp"-->
<!--#include file="../inc/ckfinder/ckfinder.asp"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
Const MAX_SMSCONTENT_LENGTH=150
Dim sendtype,tid
sendtype=Request("type")
batch=Request.Form("batch")
curstep=Request.QueryString("step")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
If Len(sendtype)=0 Or Not IsNumeric(sendtype) Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">参数错误！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End()
End If
If IsEmpty(batch) Then batch=0

Select Case curstep
Case vbNullString
	tid=Request("tid")
	If IsEmpty(tid) Then tid=Request("sel")
	If IsEmpty(tid) Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">请选择要通知的专家！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		Response.End()
	End If
	Connect conn
	sql="SELECT TEACHERNAME,MOBILE,EMAIL FROM ViewTeacherInfo WHERE TEACHERID IN ("&tid&")"
	GetRecordSetNoLock conn,rs,sql,result
	If result=1 Then
		title="给["&rs("TEACHERNAME")&"]老师"
		default_content="尊敬的"&rs("TEACHERNAME")&"老师：<br/>您好！"
	ElseIf result>1 Then
		title="给["&rs("TEACHERNAME")&"]等&nbsp;"&result&"&nbsp;名老师"
		default_content="尊敬的老师：<br/>您好！"
	End If
	If sendtype=1 Then
		title=title&"发送短信"
	Else
		title=title&"发送邮件"
	End If
	If batch=1 Then
		cancelUrl="history.go(-1)"
	Else
		cancelUrl="window.close()"
	End If
	Do While Not rs.EOF
		If Len(mobile) Then mobile=mobile&","
		mobile=mobile&rs("MOBILE")
		If Len(email) Then email=email&","
		email=email&rs("EMAIL")
		rs.MoveNext()
	Loop
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>送审结果统计表</title>
<% useStylesheet("admin") %>
<% useScript(Array("jquery", "common")) %>
</head>
<body bgcolor="ghostwhite">
<center>
<font size=4><b><%=title%></b></font>
<form id="fmSend" action="?step=2" method="post">
<table width="1000" cellpadding="2" cellspacing="1" bgcolor="dimgray"><%
	If sendtype=1 Then %>
<tr bgcolor="gainsboro" align="center" height="25">
<td><span class="tip">当前编辑的是手机短信，字数限制在<%=MAX_SMSCONTENT_LENGTH%>字以内，只能发送纯文字，暂不支持使用样式或插入图片！</span></td></tr>
<tr bgcolor="gainsboro" height="25">
<td>收件人：<input type="text" name="rcpt" size="100" value="<%=mobile%>" readonly /></td></tr><%
	ElseIf sendtype=2 Then %>
<tr bgcolor="gainsboro" height="25">
<td>收件人：<input type="text" name="rcpt" size="100" value="<%=email%>" readonly /></td></tr>
<tr bgcolor="gainsboro" height="25">
<td>标题：&emsp;<input type="text" name="subject" size="100" /></td></tr><%
	End If %>
<tr bgcolor="gainsboro" height="25">
<td>内容：<br/><% SetEditorWithName "content",default_content,130 %></td></tr>
<tr bgcolor="gainsboro" align="center" height="25">
<td><input type="submit" name="btnsubmit" value="发 送" />&emsp;<input type="button" value="取 消" onclick="<%=cancelUrl%>" /></td></tr></table>
<input type="hidden" name="finalFilter2" value="<%=finalFilter%>" />
<input type="hidden" name="pageSize2" value=<%=pageSize%> />
<input type="hidden" name="pageNo2" value=<%=pageNo%> />
<input type="hidden" name="type" value="<%=sendtype%>" />
<input type="hidden" name="tid" value="<%=tid%>" />
<input type="hidden" name="batch" value="<%=batch%>" /></form></center>
<script type="text/javascript">
	$(':submit').click(function() {
		this.value="正在发送，请稍候……";
		this.disabled=true;
	}).attr('disabled',false);
</script></body></html><%
Case 2
	subject=Request.Form("subject")
	content=Request.Form("content")
	rcpt=Request.Form("rcpt")
	If Len(content)=0 Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">请填写内容！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		Response.End()
	ElseIf sendtype=1 Then
		content=toPlainText(content)
		If Len(content)>MAX_SMSCONTENT_LENGTH Then
			%><body bgcolor="ghostwhite"><center><font color=red size="4">短信内容不能超过<%=MAX_SMSCONTENT_LENGTH%>字，请缩减后再发送或分段发送！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
			Response.End()
		End If
	ElseIf sendtype=2 Then
		If Len(subject)=0 Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">请填写邮件标题！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
			Response.End()
		End If
	End If
	
	Dim ret,bSuccess,msg,arrRcpt,numRcpt,numSuccess,i
	arrRcpt=Split(rcpt,",")
	numRcpt=UBound(arrRcpt)+1
	numSuccess=0
	For i=0 To numRcpt-1
		If sendtype=1 Then
			ret=sendCustomSMS(arrRcpt(i),content)
			bSuccess=ret=0
		ElseIf sendtype=2 Then
			bSuccess=sendCustomEmail(arrRcpt(i),subject,content)
		End If
		If Len(msg) Then msg=msg&"\n"
		If bSuccess Then
			numSuccess=numSuccess+1
			msg=msg&arrRcpt(i)&"：发送成功。"
		Else
			msg=msg&arrRcpt(i)&"：发送失败。"
		End If
		If sendtype=1 Then
			msg=msg&"(响应值："&ret&")"
		End If
	Next
	msg="发送数："&numRcpt&"，成功数："&numSuccess&"，详情如下：\n"&msg
%><form id="ret" action="expertList.asp" method="post">
<input type="hidden" name="finalFilter2" value="<%=finalFilter%>" />
<input type="hidden" name="pageSize2" value=<%=pageSize%> />
<input type="hidden" name="pageNo2" value=<%=pageNo%> /></form>
<script type="text/javascript">
	alert("<%=msg%>");<%
	If batch=1 Then %>
	document.all.ret.submit();<%
	Else %>
	window.close();<%
	End If %>
</script><%
End Select
CloseRs rs
CloseConn conn
%>
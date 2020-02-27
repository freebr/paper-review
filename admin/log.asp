<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
logDate=Request.Form("logdate")
If Len(logDate)=0 Then
	arr=Split(Date,"/")
	logDate=Format("{0}-{1}-{2}",arr(0),Right("0"&arr(1),2),Right("0"&arr(2),2))
	filename=FormatDateTime(Date,1)
Else
	arr=Split(logDate,"-")
	filename=Format("{0}年{1}月{2}日",Int(arr(0)),Int(arr(1)),Int(arr(2)))
End If
Set fso=Server.CreateObject("Scripting.FileSystemObject")
logFile=Server.MapPath("/log/PaperReview/"&filename&".log")
If fso.FileExists(logFile) Then
	Set stream=fso.OpenTextFile(logFile)
Else
	bNotExist=True
End If
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="theme-color" content="#2D79B2" />
<title>用户操作日志(<%=filename%>)</title>
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body bgcolor="ghostwhite">
<center><br>
<font size="3"><b>用户操作日志(<%=filename%>)</b></font>
<form name="fmViewlog" method="post">
<table width="400" border="0" cellspacing="1" cellpadding="3" bgcolor="gainsboro">
	<tr bgcolor="#ffffff">
		<td width="70">输入日期：</td>
		<td align="center"><input type="date" size="15" name="logdate" style="text-align:center" value="<%=logDate%>" /></td>
		<td align="center"><input type="submit" name="btnsubmit" value="确定" /></td>
	</tr>
</table></form>
<%
If bNotExist Then
%>没有该日期的日志文件！<%
Else %>
<table width="600" border=0 cellspacing=1 cellpadding=3 bgcolor="#999999" style="margin:10px 0"><%
lineStart="<tr bgcolor=""#ffffff""><td>"
lineEnd="</td></tr>"
Do While Not stream.AtEndOfStream
	logLine=stream.ReadLine()
	logContent=lineStart&logLine&lineEnd&logContent
Loop
stream.Close()
Set stream=Nothing
Set fso=Nothing
Response.Write logContent %>
</table><p align="center"><span style="text-decoration:line-through"><%=Replace(String(10," ")," ","&nbsp;")%></span>文件头<span style="text-decoration:line-through"><%=Replace(String(10," ")," ","&nbsp;")%></span></p><%
End If %>
</center>
</body>
</html>
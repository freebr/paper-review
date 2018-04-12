<%If IsEmpty(Session("Id")) Then Response.Redirect("../err/timeout.asp")
logDate=Request.Form("logdate")
If Len(logDate)=0 Then
	logDate=FormatDateTime(Date,1)
End If
Set fso=Server.CreateObject("Scripting.FileSystemObject")
logFile=Server.MapPath("/log/ThesisReview/"&logDate&".log")
If fso.FileExists(logFile) Then
	Set stream=fso.OpenTextFile(logFile)
Else
	bNotExist=True
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../css/global.css" rel="stylesheet" type="text/css">
<script src="/index/javascript/jquery-1.6.3.min.js" type="text/javascript"></script>
<meta name="theme-color" content="#2D79B2" />
<title>用户操作日志(<%=logDate%>)</title>
</head>
<body bgcolor="ghostwhite">
<center><br>
<font size="3"><strong>用户操作日志(<%=logDate%>)</strong></font>
<form name="fmViewlog" method="post">
<table width="400" border="0" cellspacing="1" cellpadding="3" bgcolor="gainsboro">
	<tr bgcolor="#ffffff">
		<td width="70">输入日期：</td>
		<td align="center"><input type="text" size="15" name="logdate" style="text-align:center" value="<%=logDate%>" /></td>
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
<%Response.Charset="utf-8"
Response.Expires=-1%>
<!-- #include File="../inc/db.asp" -->
<%If IsEmpty(Session("Tuser")) Then Response.Redirect("../error.asp?timeout")

newpwd=Request.Form("newPwd")
repeatpwd=Request.Form("repeatPwd")
Connect conn
str="SELECT * FROM TEACHER_INFO WHERE TEACHERID="&Session("Tid")
GetRecordSet conn,rs,str,result
If Len(newpwd)=0 Then
	bError=True
	errdesc="密码不能为空！"
ElseIf newpwd<>repeatpwd Then
	bError=True
	errdesc="两次输入的密码不相同！"
ElseIf rs.EOF Then
	bError=True
	errdesc="教师不存在！"
End If
If bError Then
%><html><head><link href="../css/teacher.css" rel="stylesheet" type="text/css" /></head>
<body class="exp"><center><div class="content"><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></div></body></html><%
	CloseRs rs
  CloseConn conn
	Response.End
End If
sql="UPDATE TEACHER_INFO SET USER_PASSWORD="&toSqlString(newpwd)&" WHERE TEACHERID="&Session("Tid")&" AND VALID=0"
conn.Execute sql
CloseConn conn
CloseRs rs
%><script type="text/javascript">
	alert("修改密码成功！");
	location.href="/teacher/mainF_exp.asp";
</script>
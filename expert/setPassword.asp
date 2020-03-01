<%Response.Expires=-1%>
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")

new_password=Request.Form("newPwd")
repeat_password=Request.Form("repeatPwd")
new_security_level=verifyPasswordSecurityLevel(new_password)

ConnectOriginDb conn
sql="SELECT * FROM TEACHER_INFO WHERE TEACHERID="&Session("TId")
GetRecordSet conn,rs,sql,count
If Len(new_password)=0 Then
	bError=True
	errdesc="密码不能为空！"
ElseIf new_password<>repeat_password Then
	bError=True
	errdesc="两次输入的密码不相同！"
ElseIf rs.EOF Then
	bError=True
	errdesc="教师不存在！"
ElseIf new_security_level<2 Then
	bError=True
	errdesc="密码强度不够，请确认是否满足以下要求："&vbNewLine&"长度8~24位，必须包含数字、大写字母、小写字母、特殊字符中至少三种。"
End If
If bError Then
%><html><head><% useStylesheet "tutor" %></head>
<body class="exp"><center><div class="content"><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></div></body></html><%
	CloseRs rs
	CloseConn conn
	Response.End()
End If
sql="UPDATE TEACHER_INFO SET USER_PASSWORD="&toSqlString(new_password)&" WHERE TEACHERID="&Session("TId")&" AND VALID=0"
conn.Execute sql
CloseConn conn
CloseRs rs
%><script type="text/javascript">
	alert("修改密码成功！");
	location.href="home.asp";
</script>
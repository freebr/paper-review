<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>登录密码修改</title>
<% useStylesheet "tutor" %>
<% useScript "jquery" %>
<script type="text/javascript"><%
    If Not IsEmpty(Session("PasswordWarning")) Then
%>alert("<%=toJsString(Session("PasswordWarning"))%>");<%
    End If %>
function chkInput(){
    if(document.all.newPwd.value==""){
        alert("请输入新密码!");
        document.all.newPwd.focus();
        return false;
    }
    if(document.all.repeatPwd.value==""){
        alert("请输入确认新密码!");
        document.all.repeatPwd.focus();
        return false;
    }
    if(document.all.newPwd.value!=document.all.repeatPwd.value){
        alert("新密码和确认新密码不一致!");
        document.all.repeatPwd.focus();
        return false;
    }
    return true;
}
</script>
</head>
<body class="exp"><center><div class="content">
<form action="setPassword.asp" method="post" onsubmit="return chkInput()">
<input type="hidden" name="teacherid" value="<%=TeacherId%>">
<h2><b><font class="title">登录密码修改</font></b></h2>
<table class="form" width="750" cellspacing="1" cellpadding="3">
<tr>
<td colspan="3" align="center">
您当前的密码强度：<span class="pwd-security-level-<%=Session("PasswordSecurityLevel")%>">【<%=arrSecurityLevelName(Session("PasswordSecurityLevel"))%>】</span>
</td>
</tr>
<tr>
<td>
新密码
</td>
<td>
<input type="password" name="newPwd" style="width:150px">
</td>
<td align="center">
请输入新密码<span style="color: red">（长度8~24位，必须包含数字、大写字母、小写字母、特殊字符中至少三种）</span>
</td>
</tr>
<tr>
<td>
确认新密码
</td>
<td>
<input type="password" name="repeatPwd" style="width:150px">
</td>
<td align="center">
请再输入一次新密码
</td>
</tr>
<tr>
<td colspan="3"><p align="center"><input type="submit" value="确认修改" /></p></td>
</tr>
</table>
</form>
</div>
</center>
</body>
</html>
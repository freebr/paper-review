<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
TeacherId=Request.QueryString("id")
If Len(TeacherId)=0 Or Not IsNumeric(TeacherId) Then
	bError=True
	errMsg="参数无效。"
End If
If bError Then
	CloseRs rs
	CloseConn conn
	showErrorPage errMsg, "提示"
End If

Connect conn
sql="SELECT * FROM ViewExpertInfo WHERE TEACHER_ID="&TeacherId
GetRecordSetNoLock conn,rs,sql,count
If rs.EOF Then
	CloseRs rs
	CloseConn conn
	showErrorPage "数据库没有记录！", "提示"
End If
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>查看评阅专家信息-<%=rs("EXPERT_NAME")%></title>
<% useStylesheet "admin" %>
<% useScript "jquery", "common", "expert" %>
</head>
<body><center>
<form id="profile" action="updateExpProfile.asp" method="post" enctype="multipart/form-data">
<input type="hidden" name="teacherid" value="<%=TeacherId%>">
<span class="title">查看评阅专家信息</span>
<p><span class="tip">以下信息均为必填项</span></p>
<table class="form profile" width="1000" cellspacing="1" cellpadding="3">
<tr height="30">
	<td bgcolor="gainsboro" align="center">姓名</td>
	<td bgcolor="white"><input type="text" name="teachername" class="txt full-width" value="<%=HtmlEncode(rs("EXPERT_NAME").Value)%>" /></td>
	<td bgcolor="gainsboro" align="center">性别</td>
	<td bgcolor="white"><input type="radio" name="sex" value="男"<%If rs("SEX").Value="男" Then%> checked<%End If%> />男
	<input type="radio" name="sex" value="女"<%If rs("SEX").Value="女" Then%> checked<%End If%> />女</td>
	<td bgcolor="gainsboro" align="center">专业技术职务</td>
	<td bgcolor="white"><input type="text" class="txt full-width" name="pro_duty_name" id="pro_duty_name" value="<%=HtmlEncode(rs("PRO_DUTY_NAME").Value)%>" /></td>
</tr>
<tr height="30">
	<td bgcolor="gainsboro" align="center">学科专长</td>
	<td bgcolor="white"><input type="text" class="txt full-width" name="expertise" id="expertise" value="<%=HtmlEncode(rs("EXPERTISE").Value)%>" /></td>
	<td bgcolor="gainsboro" align="center">电子邮箱</td>
	<td bgcolor="white"><input name="email" type="text" class="txt full-width" id="email" value="<%=HtmlEncode(rs("EMAIL").Value)%>" /></td>
	<td bgcolor="gainsboro" align="center">邮编</td>
	<td bgcolor="white"><input type="text" class="txt full-width" name="mailcode" id="mailcode" value="<%=HtmlEncode(rs("MAILCODE").Value)%>" onkeyup="replNoNum(this)" onbeforepaste="replClipboardData('trimNoNum')" /></td>
</tr>
<tr height="30">
	<td bgcolor="gainsboro" align="center">联系电话（办公室）</td>
	<td bgcolor="white"><input type="text" class="txt full-width" name="telephone" id="telephone" value="<%=HtmlEncode(rs("TELEPHONE").Value)%>" /></td>
	<td bgcolor="gainsboro" align="center">联系电话（移动）</td>
	<td bgcolor="white"><input type="text" class="txt full-width" name="mobile" id="mobile" value="<%=HtmlEncode(rs("MOBILE").Value)%>" onkeyup="replNoNum(this)" onbeforepaste="replClipboardData('trimNoNum')" /></td>
	<td bgcolor="gainsboro" align="center">最高学历</td>
	<td bgcolor="white"><%=diplomaList("last_diploma",rs("LAST_DIPLOMA").Value)%></td>
</tr>
<tr height="30">
	<td bgcolor="gainsboro" align="center">登录名</td>
	<td bgcolor="white"><%
	If rs("INSCHOOL").Value Then
%><%=rs("TEACHERNO").Value%><%
	Else
%><input type="text" class="txt full-width" name="teacherno" value="<%=rs("TEACHERNO").Value%>" /><%
	End If
%></td>
	<td bgcolor="gainsboro" align="center">身份证号码</td>
	<td bgcolor="white" colspan="5"><input type="text" class="txt full-width" name="idcard_no" id="idcard_no" value="<%=HtmlEncode(rs("IDCARD_NO").Value)%>" /></td>
</tr>
<tr height="30">
	<td bgcolor="gainsboro" align="center">银行账户号</td>
	<td bgcolor="white"><input type="text" class="txt full-width" name="bankaccount" id="bankaccount" value="<%=HtmlEncode(rs("BANK_ACCOUNT").Value)%>" onkeyup="replNoNum(this)" onbeforepaste="replClipboardData('trimNoNum')" /></td>
	<td bgcolor="gainsboro" align="center">开户行</td>
	<td bgcolor="white" colspan="3"><input type="text" class="txt full-width" name="bankname" id="bankname" value="<%=HtmlEncode(rs("BANK_NAME").Value)%>" /></td>
</tr>
<tr height="30">
	<td bgcolor="gainsboro" align="center">单位名称（含院系）</td>
	<td bgcolor="white" colspan="5"><input type="text" class="txt full-width" name="workplace" id="workplace" value="<%=HtmlEncode(rs("WORKPLACE").Value)%>" /></td>
</tr>
<tr height="30">
	<td bgcolor="gainsboro" align="center">通信地址（最多25字）</td>
	<td bgcolor="white" colspan="5"><input type="text" class="txt full-width" name="address" id="address" value="<%=HtmlEncode(rs("ADDRESS").Value)%>" /></td>
</tr>
<tr bgcolor="white"><td colspan="6" align="center">
<input type="button" id="btnsubmit" value="提 交" onclick="submitForm(this.form)" />&nbsp;
<input type="button" id="btnreturn" value="关 闭" onclick="closeWindow()" /></td></tr></table>
<span class="title">登录密码修改</span>
<table class="form" width="1000" cellspacing="1" cellpadding="3">
<tr height="30">
<td bgcolor="white" colspan="6" align="center">请输入新密码：<input type="password" class="txt full-width" name="newpwd" id="newpwd" style="width:150px" />&emsp;
确认新密码：<input type="password" class="txt full-width" name="repeatpwd" id="repeatpwd" style="width:150px" /></td>
</tr>
<tr bgcolor="white"><td colspan="6" align="center">
<input type="button" value="修改密码" onclick="submitForm(this.form)" /></td></tr></table></form></center></body></html><%
CloseRs rs
CloseConn conn
%>
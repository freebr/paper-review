<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("Tid")) Then Response.Redirect("../error.asp?timeout")
TeacherId=Session("Tid")
If Len(TeacherId)=0 Or Not IsNumeric(TeacherId) Then
	bError=True
	errdesc="参数无效。"
End If
If bError Then
%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	CloseRs rs
	CloseConn conn
	Response.End()
End If

Connect conn
sql="SELECT * FROM ViewExpertInfo WHERE TEACHER_ID="&TeacherId
GetRecordSetNoLock conn,rs,sql,result
If rs.EOF Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">数据库没有记录！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	CloseRs rs
  CloseConn conn
	Response.End()
End If
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<% useStylesheet("tutor") %>
<% useScript(Array("jquery", "common", "expert")) %>
<title>个人信息编辑</title>
</head>
<body class="exp"><center><div class="content">
<form id="profile" action="updateProfile.asp" method="post" enctype="multipart/form-data">
<span class="title">个人信息编辑</span>
<p><span class="tip">以下信息均为必填项</span></p>
<table class="tblform profile" width="1000" cellspacing="1" cellpadding="3">
<tr height="30">
	<td bgcolor="gainsboro" align="center">姓名</td>
	<td bgcolor="white"><input type="text" name="teachername" class="txt" value="<%=HtmlEncode(rs("EXPERT_NAME").Value)%>" /></td>
	<td bgcolor="gainsboro" align="center">性别</td>
	<td bgcolor="white"><input type="radio" name="sex" value="男"<%If rs("SEX").Value="男" Then%> checked<%End If%> />男
	<input type="radio" name="sex" value="女"<%If rs("SEX").Value="女" Then%> checked<%End If%> />女</td>
	<td bgcolor="gainsboro" align="center">专业技术职务</td>
	<td bgcolor="white"><input type="text" class="txt" name="pro_duty_name" id="pro_duty_name" value="<%=HtmlEncode(rs("PRO_DUTY_NAME").Value)%>" /></td>
</tr>
<tr height="30">
	<td bgcolor="gainsboro" align="center">学科专长</td>
	<td bgcolor="white"><input type="text" class="txt" name="expertise" id="expertise" value="<%=HtmlEncode(rs("EXPERTISE").Value)%>" /></td>
	<td bgcolor="gainsboro" align="center">电子邮箱</td>
	<td bgcolor="white"><input name="email" type="text" class="txt" id="email" value="<%=HtmlEncode(rs("EMAIL").Value)%>" /></td>
	<td bgcolor="gainsboro" align="center">邮编</td>
	<td bgcolor="white"><input type="text" class="txt" name="mailcode" id="mailcode" value="<%=HtmlEncode(rs("MAILCODE").Value)%>" onkeyup="replNoNum(this)" onbeforepaste="replClipboardData('trimNoNum')" /></td>
</tr>
<tr height="30">
	<td bgcolor="gainsboro" align="center">联系电话（办公室）</td>
	<td bgcolor="white"><input type="text" class="txt" name="telephone" id="telephone" value="<%=HtmlEncode(rs("TELEPHONE").Value)%>" /></td>
	<td bgcolor="gainsboro" align="center">联系电话（移动）</td>
	<td bgcolor="white"><input type="text" class="txt" name="mobile" id="mobile" value="<%=HtmlEncode(rs("MOBILE").Value)%>" onkeyup="replNoNum(this)" onbeforepaste="replClipboardData('trimNoNum')" /></td>
	<td bgcolor="gainsboro" align="center">最高学历</td>
	<td bgcolor="white"><%=diplomaList("last_diploma",rs("LAST_DIPLOMA").Value)%></td>
</tr>
<tr height="30">
	<td bgcolor="gainsboro" align="center">身份证号码</td>
	<td bgcolor="white" colspan="5"><input type="text" class="txt" name="idcard_no" id="idcard_no" value="<%=HtmlEncode(rs("IDCARD_NO").Value)%>" /></td>
</tr>
<tr height="30">
	<td bgcolor="gainsboro" align="center">银行账户号</td>
	<td bgcolor="white"><input type="text" class="txt" name="bankaccount" id="bankaccount" value="<%=HtmlEncode(rs("BANK_ACCOUNT").Value)%>" onkeyup="replNoNum(this)" onbeforepaste="replClipboardData('trimNoNum')" />
	<span class="tip">请填写工行账号</span></td>
	<td bgcolor="gainsboro" align="center">开户行</td>
	<td bgcolor="white" colspan="3"><input type="text" class="txt" name="bankname" id="bankname" value="<%=HtmlEncode(rs("BANK_NAME").Value)%>" /></td>
</tr>
<tr height="30">
	<td bgcolor="gainsboro" align="center">单位名称（含院系）</td>
	<td bgcolor="white" colspan="5"><input type="text" class="txt" name="workplace" id="workplace" value="<%=HtmlEncode(rs("WORKPLACE").Value)%>" /></td>
</tr>
<tr height="30">
	<td bgcolor="gainsboro" align="center">通信地址（最多25字）</td>
	<td bgcolor="white" colspan="5"><input type="text" class="txt" name="address" id="address" value="<%=HtmlEncode(rs("ADDRESS").Value)%>" /></td>
</tr>
<tr bgcolor="white"><td colspan="6" align="center">
<input type="button" id="btnsubmit" value="提 交" />&nbsp;
<input type="button" id="btnreturn" value="返 回" /></td></tr></table>
<span class="title">登录密码修改</span>
<table class="tblform" width="1000" cellspacing="1" cellpadding="3">
<tr height="30">
<td bgcolor="white" colspan="6" align="center">请输入新密码：<input type="password" class="txt" name="newpwd" id="newpwd" style="width:150px" />&emsp;
确认新密码：<input type="password" class="txt" name="repeatpwd" id="repeatpwd" style="width:150px" /></td>
</tr>
<tr bgcolor="white"><td colspan="6" align="center">
<input type="button" value="修改密码" onclick="submitForm(this.form)" /></td></tr></table></form></div></center></body></html><%
CloseRs rs
CloseConn conn
%>
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>系统首页</title>
<% useStylesheet "tutor" %>
<% useScript "jquery" %>
</head>
<body class="exp">
	<center>
		<div class="content">
			<p><span class="maintitle">欢迎使用<br/>工商管理学院专业学位论文评阅系统</span></p>
			<p align="left">&emsp;&emsp;评阅论文前，请先在左边菜单选择<b>“个人信息编辑”</b>，按提示填写个人信息；然后在<b>“浏览论文列表”</b>处点击相应的论文条目评阅论文。</p>
			<p align="left" style="color:red">&emsp;&emsp;请注意：务必在个人信息中填写您的工商银行账号，并核对账号是否正确，如非工行账号，请及时修改。</p>
		</div>
	</center>
</body>
</html>
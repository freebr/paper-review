<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"-->
<%Response.Expires=-1
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
stuType=Request.Form("stuType")
If IsEmpty(stuType) Or Not IsNumeric(stuType) Then
	stuType=0
End If

Dim dictNotices:Set dictNotices=Server.CreateObject("Scripting.Dictionary")

Dim conn,rs,sql,num
json=Request.Form("json")
If json="1" Then
	Dim ret
	sql="EXEC getNoticeText ?"
	Set rs=ExecQuery(conn,sql,CmdParam("StudentType",adInteger,adParamInput,4,stuType),num)
	Do While Not rs.EOF
		dictNotices.Add rs(0).Value, rs(1).Value
		rs.MoveNext()
	Loop
	If num=0 Then
%>{"status": "empty"}<%
	Else
		ret="{""status"": ""ok"", ""notices"": ["
		Dim keys:keys=dictNotices.Keys()
		Dim items:items=dictNotices.Items()
		For i=0 To dictNotices.Count-1
			If i>0 Then ret=ret&","
			ret=ret&"{""name"": """&toJsString(keys(i))&""", ""content"": """&toJsString(items(i))&"""}"
		Next
		ret=ret&"]}"
		Response.Write ret
	End If
	CloseRs rs
	CloseConn conn
	Response.End()
Else
	dictNotices.Add "review_eval_reference", "送审评语的基本内容参考"
	dictNotices.Add "review_result_desc", "论文检测结果及论文评审结果说明"
End If
curstep=Request.QueryString("step")
If Len(curstep)=0 Or Not IsNumeric(curstep) Then curstep="1"

Select Case curstep
Case "1"
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
<script src="../scripts/query.js" type="text/javascript"></script>
<script src="../scripts/utils.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>提示文本设置</b></font>
<form id="fmSettings" method="post" action="?step=2">
<table width="900" cellpadding="2" cellspacing="1" bgcolor="dimgray">
<tr bgcolor="ghostwhite">
<td align="left"><%
	Select Case Request("ok")
	Case "0"
%><span style="color:red">更新设置时发生错误，请检查设置参数是否正确</span><br/><%
	Case "1"
%><span style="color:red">设置成功</span><br/><%
	End Select
%></td></tr>
<tr bgcolor="ghostwhite">
<td>学生类型：<select id="stuType" name="stuType">
	<option value="0">【请选择】</option><%
	For i=1 To UBound(arrStuTypeName) %>
	<option value="<%=arrStuTypeId(i)%>"><%=arrStuTypeName(i)%></option><%
	Next %>
</select></td></tr>
<tr bgcolor="ghostwhite">
<td><table class="notice-list" width="100%" cellpadding="0" cellspacing="0" bgcolor="ghostwhite"><%
	For i=0 To dictNotices.Count-1
%><tr bgcolor="ghostwhite"><td align="right"><%=dictNotices.Items()(i)%>：</td><td align="center">
	<textarea class="notice-text" name="<%=dictNotices.Keys()(i)%>"></textarea>
</td></tr><%
	Next %>
</table></td></tr>
<tr bgcolor="ghostwhite"><td align="center"><input type="submit" value="更改设置">&emsp;<input type="button" id="btnreturn" value="返回系统设置首页"></td></tr>
</table></form>
</center></body>
<script type="text/javascript">
	$("select#stuType").change(function() {
		$.ajax({url: 'noticeText.asp', type: 'post', data: {stuType: $(this).val(), json: 1}, dataType: 'json',
						success: function(data, status) {
							$("textarea").val('');
							if(data.status==='ok') {
								for(i in data.notices) {
									$("textarea[name='"+data.notices[i].name+"']").val(data.notices[i].content);
								}
							}
						}
					});
	})<%
	If stuType<>0 Then %>
	.val(<%=stuType%>).change();<%
	End If %>
	$("input#btnreturn").click(function() {
		location.href="systemSettings.asp";
	});
</script></html><%
Case "2"
	Dim ok:ok="1"
	Dim paramStudentType:Set paramStudentType=CmdParam("StudentType",adInteger,adParamInput,4,stuType)
	Connect conn
	For i=0 To dictNotices.Count-1
		Dim key:key=dictNotices.Keys()(i)
		Dim content:content=Request.Form(key)
		If IsEmpty(content) Then content=""
		setNoticeText stuType,key,content
	Next
	
	CloseConn conn
%><body><form id="ret" method="post" action="?step=1"><input type="hidden" name="ok" value="<%=ok%>" />
<input type="hidden" name="stuType" value="<%=stuType%>">
</form>
<script type="text/javascript">
document.all.ret.submit();
</script></body><%
End Select
%>
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%Response.Expires=-1
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
userId=Request.Form("userId")
If IsEmpty(userId) Or Not IsNumeric(userId) Then
	userId=0
End If

Dim conn,rs,sql,num
json=Request.Form("json")
If json="1" Then
	Dim ret
	Set dict=getAdminType(userId)
	
	If dict.Items()(0)=0 Then
%>{"status": "empty"}<%
	Else
		ret="{""status"": ""ok"", ""admin_type"": "&dict.Items()(0)&"}"
		Response.Clear()
		Response.Write ret
	End If
	CloseRs rs
	CloseConn conn
	Response.End()
Else
	sql="SELECT * FROM ViewAdminUsers WHERE AdminType IS NOT NULL ORDER BY UserName;" &_
		"SELECT * FROM ViewAdminUsers WHERE AdminType IS NULL ORDER BY UserName"
	GetRecordSetNoLock conn,rs,sql,count
End If
step=Request.QueryString("step")
If Len(step)=0 Or Not IsNumeric(step) Then step="1"

Select Case step
Case "1"
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body>
<center><font size=4><b>教务员类型设置</b></font>
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
<td>教务员：<select name="userId" id="userId"><option value="0" hidden>请选择</option><%
	If Not rs.EOF Then %>
	<option class="user-id-category" value="" disabled>已设置类型的教务员</option><%
		Do While Not rs.EOF %>
	<option value="<%=rs("UserID").Value%>"><%=rs("UserName").Value%>(<%=rs("LoginName").Value%>)</option><%
			rs.MoveNext()
		Loop
	End If
	Set rs = rs.NextRecordSet()
	If Not rs.EOF Then %>
	<option class="user-id-category" value="" disabled>未设置类型的教务员</option><%
		Do While Not rs.EOF %>
	<option value="<%=rs("UserID").Value%>"><%=rs("UserName").Value%>(<%=rs("LoginName").Value%>)</option><%
			rs.MoveNext()
		Loop
	End If %>
</select></td></tr>
<tr bgcolor="ghostwhite">
<td>分管学生类型：<%
	For Each item In dictStuTypes %>
	<label for="manage_stu_type<%=item%>">
	<input type="checkbox" name="manage_stu_type" id="manage_stu_type<%=item%>" value="<%=item%>" /><%=dictStuTypes(item)(1)%>
	</label><%
	Next %>
</td></tr>
<tr bgcolor="ghostwhite"><td align="center"><input type="submit" value="更改设置"></td></tr>
</table></form>
</center></body>
<script type="text/javascript">
	$("select#userId").change(function() {
		$.ajax({url: 'admins.asp', type: 'post', data: {userId: $(this).val(), json: 1}, dataType: 'json',
			success: function(data, status) {
				$("[name='manage_stu_type']").each(function(index, item) {
					item.checked=(data.admin_type&Math.pow(2,item.value-1))!==0;
				});
			}
		});
	})<%
	If userId<>0 Then %>
	.val(<%=userId%>).change();<%
	End If %>
	$(":submit").click(function() {
		if($("#userId").val()==="0") {
			alert("请选择要修改类型的教务员！");
			return false;
		}
	});
</script></html><%
Case "2"
	Dim ok:ok="1"
	Dim arrManageStuTypes:arrManageStuTypes=Split(Request.Form("manage_stu_type")+"",",")
	setAdminType userId, arrManageStuTypes
%><body><form id="ret" method="post" action="?step=1"><input type="hidden" name="ok" value="<%=ok%>" />
<input type="hidden" name="userId" value="<%=userId%>">
</form>
<script type="text/javascript">
	document.all.ret.submit();
</script></body><%
End Select
%>
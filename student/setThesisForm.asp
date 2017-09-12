<%Response.Charset="utf-8"
Response.Expires=-1%>
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("user")) Then Response.Redirect("../error.asp?timeout")
thesisID=Request.Form("tid")
new_thesis_form=Request.Form("thesis_form")
If IsEmpty(thesisID) Then
	bError=True
	errdesc="参数无效。"
ElseIf new_thesis_form=0 Then
	bError=True
	errdesc="请选择论文形式！"
End If
If bError Then
%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End
End If

Dim conn,sql
Connect conn
sql="UPDATE TEST_THESIS_REVIEW_INFO SET REVIEW_TYPE="&new_thesis_form&" WHERE ID="&thesisID
conn.Execute sql
CloseConn conn
%><script type="text/javascript">
	alert("操作完成。");
	location.href='default.asp';
</script>
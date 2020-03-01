<!--#include file="../inc/global.inc"-->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
Dim ids
ids=Request.Form("sel")
If Len(ids)=0 Then
	showErrorPage "请选择要删除的论文记录！", "提示"
End If

FormGetToSafeRequest(ids)
Connect conn
sql="DELETE FROM Dissertations WHERE ID IN ("&ids&")"
conn.Execute sql
CloseConn conn
%><script type="text/javascript">
	alert("操作完成。");
	location.href="paperList.asp";
</script>
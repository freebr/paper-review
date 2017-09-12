<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<%If IsEmpty(Session("user")) Then Response.Redirect("../error.asp?timeout")
Dim ids
ids=Request.Form("sel")
FormGetToSafeRequest(ids)
Connect conn
sql="DELETE FROM TEST_THESIS_REVIEW_INFO WHERE ID IN ("&ids&")"
conn.Execute sql
CloseConn conn
%><script type="text/javascript">
	alert("操作完成。");
	location.href="thesisList.asp";
</script>
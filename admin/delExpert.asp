<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
Dim ids,sel_count
ids=Request.Form("sel")
sel_count=Request.Form("sel").Count
FormGetToSafeRequest(ids)
Connect conn
ConnectOriginDb connOrigin
For i=1 To sel_count
	sql_origin=sql_origin&"UPDATE TEACHER_INFO SET WRITEPRIVILEGETAGSTRING=dbo.removePrivilege(WRITEPRIVILEGETAGSTRING,'I10'),"&_
				  "READPRIVILEGETAGSTRING=dbo.removePrivilege(READPRIVILEGETAGSTRING,'I10') WHERE TEACHERID="&Request.Form("sel")(i)&";"
Next
If sel_count>0 Then
	sql=sql&"DELETE FROM Experts WHERE TEACHER_ID IN ("&ids&");"
	sql_origin=sql_origin&"DELETE FROM TEACHER_INFO WHERE TEACHERID IN ("&ids&") AND IFTEACHER=3;"
	conn.Execute sql
	connOrigin.Execute sql_origin
End If
CloseConn connOrigin
CloseConn conn
%><script type="text/javascript">
	alert("操作完成。");
	location.href="expertList.asp";
</script>
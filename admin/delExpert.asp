﻿<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<%If IsEmpty(Session("user")) Then Response.Redirect("../error.asp?timeout")
Dim ids,sel_count
ids=Request.Form("sel")
sel_count=Request.Form("sel").Count
FormGetToSafeRequest(ids)
Connect conn
For i=1 To sel_count
	sql=sql&"UPDATE TEACHER_INFO SET WRITEPRIVILEGETAGSTRING=dbo.removePrivilege(WRITEPRIVILEGETAGSTRING,'I10'),"&_
				  "READPRIVILEGETAGSTRING=dbo.removePrivilege(READPRIVILEGETAGSTRING,'I10') WHERE TEACHERID="&Request.Form("sel")(i)&";"
Next
If sel_count>0 Then
	sql=sql&"DELETE FROM TEST_THESIS_REVIEW_EXPERT_INFO WHERE TEACHER_ID IN ("&ids&");"
	sql=sql&"DELETE FROM TEACHER_INFO WHERE TEACHERID IN ("&ids&") AND IFTEACHER=3;"
	conn.Execute sql
End If
CloseConn conn
%><script type="text/javascript">
	alert("操作完成。");
	location.href="expertList.asp";
</script>
﻿<%Response.Charset="utf-8"
Response.Expires=-1%>
<!--#include file="../inc/db.asp"-->
<%If IsEmpty(Session("user")) Then Response.Redirect("../error.asp?timeout")

ids=Request.Form("sel")
sel_count=Request.Form("sel").Count
FormGetToSafeRequest(ids)
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
password="123456"
ConnectOriginDb conn
sql="UPDATE TEACHER_INFO SET USER_PASSWORD="&toSqlString(password)&" WHERE TEACHERID IN ("&ids&")"
conn.Execute sql
CloseConn conn

%><form id="ret" action="expertList.asp" method="post">
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form>
<script type="text/javascript">
	alert("操作完成，已将 <%=sel_count%> 名专家的账号密码重置为 <%=toJsString(password)%>。");
	document.all.ret.submit();
</script>
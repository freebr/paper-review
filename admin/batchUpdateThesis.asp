<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
Dim reviewfilestat,ids
reviewfilestat=Request.Form("reviewfilestat")
ids=Request.Form("sel")
If Len(ids)=0 Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">请选择论文记录！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End
End If

teachtype_id=Request.Form("In_TEACHTYPE_ID2")
class_id=Request.Form("In_CLASS_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
FormGetToSafeRequest(reviewfilestat)
FormGetToSafeRequest(ids)
Connect conn
sql="UPDATE TEST_THESIS_REVIEW_INFO SET REVIEW_FILE_STATUS="&reviewfilestat&" WHERE ID IN ("&ids&")"
conn.Execute sql
CloseConn conn
%><form id="ret" action="thesisList.asp" method="post">
<input type="hidden" name="In_TEACHTYPE_ID" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize" value="<%=pageSize%>" />
<input type="hidden" name="pageNo" value="<%=pageNo%>" /></form>
<script type="text/javascript">
	alert("操作完成。");
	document.all.ret.submit();
</script>
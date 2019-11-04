<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
Dim review_display_status,ids
review_display_status=Request.Form("review_display_status")
ids=Request.Form("sel")
If Len(ids)=0 Then
	showErrorPage "请选择论文记录！", "提示"
End If

activity_id=Request.Form("In_ActivityId2")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
class_id=Request.Form("In_CLASS_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
FormGetToSafeRequest(review_display_status)
FormGetToSafeRequest(ids)
Connect conn
sql="SELECT THESIS_SUBJECT, STU_NAME FROM ViewDissertations_admin WHERE ID IN ("&ids&")"
GetRecordSetNoLock conn,rs,sql,count

Dim titles: titles=""
Do While Not rs.EOF
	If Len(titles) Then titles=titles&"；"
	titles=titles&Format("《{0}》，作者：{1}", Array(rs(0), rs(1)))
	rs.MoveNext()
Loop
CloseRs rs

sql=Format("UPDATE Dissertations SET REVIEW_FILE_STATUS={0} WHERE ID IN ({1})",_
	Array(review_display_status, ids))
conn.Execute sql

sql=Format("UPDATE ReviewRecords SET DisplayStatus={0},DisplayStatusModifiedBy={1} WHERE DissertationId IN ({2})",_
	Array(review_display_status, Session("Id"), ids))
conn.Execute sql
CloseConn conn

writeEventLog 0,Session("Name"),"对以下论文修改评阅书开放状态为["&arrReviewFileStat(review_display_status)&"]："&titles&"。"
%><form id="ret" action="thesisList.asp" method="post">
<input type="hidden" name="In_ActivityId" value="<%=activity_id%>">
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
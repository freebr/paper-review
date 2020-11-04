<%Response.Expires=-1%>
<!--#include file="../inc/automation/ReviewApplicationFormWriter.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
paper_id=Request.QueryString("tid")
activity_id=Request.Form("In_ActivityId2")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
class_id=Request.Form("In_CLASS_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
If Len(paper_id)=0 Or Not IsNumeric(paper_id) Then
	showErrorPage "参数无效。", "提示"
End If
sql="SELECT * FROM ViewDissertations WHERE ID="&paper_id
GetRecordSet conn,rs,sql,count
' 生成送审申请表
Dim rag
Randomize()
review_time=rs("SUBMIT_REVIEW_TIME")
If IsNull(review_time) Then review_time=Now
eval_text=rs("REVIEW_APP_EVAL")
If IsNull(eval_text) Then eval_text=""
filename=toDateTime(review_time,1)&Int(Timer)&Int(Rnd()*999)&".doc"
filepath=Server.MapPath(basePath()&"tutor/export/"&filename)
Set rag=New ReviewApplicationFormWriter
rag.Author=rs("STU_NAME")
rag.StuNo=rs("STU_NO")
rag.TutorInfo=rs("TUTOR_NAME")&" "&getProDutyNameOf(rs("TUTOR_ID"))
rag.Spec=rs("SPECIALITY_NAME")
rag.Date=toDateTime(review_time,1)
rag.Subject=rs("THESIS_SUBJECT")
rag.Comment=eval_text
rag.ReproductRatio=rs("REPRODUCTION_RATIO")
bError=rag.generateApp(filepath)=0
Set rag=Nothing
rs("REVIEW_APP")=filename
rs.Update()
CloseRs rs
CloseConn conn
%><form id="ret" action="paperDetail.asp?tid=<%=paper_id%>" method="post">
<input type="hidden" name="In_ActivityId2" value="<%=activity_id%>">
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID2" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form>
<script type="text/javascript">
	alert("操作完成。");
	document.all.ret.submit();
</script>
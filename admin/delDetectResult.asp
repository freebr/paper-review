<%Response.Expires=-1%>
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
paper_id=Request.Form("tid")
detect_id=Request.Form("id")
delete_type=Request.Form("delete_type")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
class_id=Request.Form("In_CLASS_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
If IsEmpty(paper_id) Or Not IsNumeric(paper_id) Or IsEmpty(detect_id) Or IsEmpty(delete_type) Or Not IsNumeric(delete_type) Or Not (delete_type=0 Or delete_type=1) Then
	showErrorPage "参数无效。", "提示"
End If

Dim conn,rs,sql,count
ConnectDb conn
sql="SELECT THESIS_FILE,THESIS_FILE4,REPRODUCTION_RATIO,INSTRUCT_REVIEW_REPRODUCTION_RATIO,DETECT_REPORT,INSTRUCT_REVIEW_DETECT_REPORT FROM Dissertations WHERE ID="&paper_id
GetRecordSet conn,rs,sql,count
If rs.EOF Then
	CloseRs rs
	CloseConn conn
	showErrorPage "数据库没有该论文记录！", "提示"
End If

sql=Format("SELECT *,dbo.isLatestDetectResult('{0}') AS IsLatest FROM DetectResults WHERE Id='{0}'",detect_id)
GetRecordSet conn,rsDetect,sql,count
detect_type=rsDetect("DetectType")
is_latest=rsDetect("IsLatest")

If delete_type=0 Then	' 删除送检报告
	rsDetect("Result")=Null
	rsDetect("DetectTime")=Null
	rsDetect("ReportFile")=Null
	rsDetect.Update()
ElseIf delete_type=1 Then	' 删除送检记录
	rsDetect.Delete()
End If
CloseRs rsDetect

If is_latest Then	' 更新论文评阅信息表中的检测数据
	Dim arrDetectFileFieldNames:arrDetectFileFieldNames=Array("","THESIS_FILE","THESIS_FILE4")
	Dim arrDetectResultFieldNames:arrDetectResultFieldNames=Array("","REPRODUCTION_RATIO","INSTRUCT_REVIEW_REPRODUCTION_RATIO")
	Dim arrDetectReportFieldNames:arrDetectReportFieldNames=Array("","DETECT_REPORT","INSTRUCT_REVIEW_DETECT_REPORT")
	sql="SELECT DetectFile,Result,ReportFile FROM DetectResults WHERE DissertationId=? AND DetectType=? ORDER BY DetectTime DESC"
	Set ret=ExecQuery(conn,sql,_
		CmdParam("DissertationId",adInteger,4,paper_id),_
		CmdParam("DetectType",adInteger,4,detect_type))
	If ret("count")>0 Then	' 取上次的送检结果
		If delete_type=1 Then rs(arrDetectFileFieldNames(detect_type))=ret("rs")("DetectFile")
		rs(arrDetectResultFieldNames(detect_type))=ret("rs")("Result")
		rs(arrDetectReportFieldNames(detect_type))=ret("rs")("ReportFile")
	Else
		If delete_type=1 Then rs(arrDetectFileFieldNames(detect_type))=Null
		rs(arrDetectResultFieldNames(detect_type))=Null
		rs(arrDetectReportFieldNames(detect_type))=Null
	End If
	rs.Update()
	CloseRs ret("rs")
End If
CloseRs rs
CloseConn conn
%><form id="ret" action="paperDetail.asp?tid=<%=paper_id%>" method="post">
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
<%Response.Charset="utf-8"
Response.Expires=-1%>
<!--#include file="appgen.inc"-->
<!--#include file="evalappend.inc"-->
<!--#include file="../inc/db.asp"-->
<!--#include file="../inc/mail.asp"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")
curstep=Request.QueryString("step")
thesisID=Request.QueryString("tid")
opr=Request.Form("opr")
submittype=Request.Form("submittype")
ispass=submittype="pass"
eval_text=Request.Form("eval_text")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
spec_id=Request.Form("In_SPECIALITY_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
If nSystemStatus<>2 Then
	bError=True
	errdesc="评阅系统的开放时间为"&toDateTime(startdate,1)&"至"&toDateTime(enddate,1)&"，当前不在开放时间内，不能审核论文。"
ElseIf Len(thesisID)=0 Or Not IsNumeric(thesisID) Then
	bError=True
	errdesc="参数无效。"
ElseIf Not isMatched("[1-8]",opr) Then
	bError=True
	errdesc="操作无效。"
ElseIf submittype="unpass" And opr<=3 Or opr=4 Or opr=5 Or opr=6 Or opr=8 Or opr=9 Then
	If Len(eval_text)=0 Then
		bError=True
		errdesc="请填写意见（200-2000字）！"
	ElseIf Len(eval_text)>2000 Then
		bError=True
		errdesc="意见字数超出限制（2000字）！"
	End If
End If
If bError Then
%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	CloseRs rs
	CloseConn conn
	Response.End
End If

Dim conn,rs,sql,result
Connect conn
sql="SELECT * FROM TEST_THESIS_REVIEW_INFO WHERE ID="&thesisID
GetRecordSet conn,rs,sql,result
review_status=rs("REVIEW_STATUS")
If rs.EOF Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">数据库没有该论文记录！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
  CloseRs rs
  CloseConn conn
	Response.End
End If

Select Case opr
Case 1	'	 审核开题报告表
	filetypename="开题报告表/开题论文"
	If ispass Then
		rs("TASK_PROGRESS")=tpTbl1Passed
		rs("TASK_EVAL")=Null
	Else
		rs("TASK_PROGRESS")=tpTbl1Unpassed
		rs("TASK_EVAL")=eval_text
	End If
	rs.Update()
	CloseRs rs
	CloseConn conn
Case 2	'  审核中期检查表
	filetypename="中期检查表/中期论文"
	If ispass Then
		rs("TASK_PROGRESS")=tpTbl2Passed
		rs("TASK_EVAL")=Null
	Else
		rs("TASK_PROGRESS")=tpTbl2Unpassed
		rs("TASK_EVAL")=eval_text
	End If
	rs.Update()
	CloseRs rs
	CloseConn conn
Case 3	'  审核预答辩申请
	filetypename="预答辩申请表/预答辩论文"
	If ispass Then
		' 更新记录
		rs("TASK_PROGRESS")=tpTbl3Passed
		rs("TASK_EVAL")=Null
	Else
		rs("TASK_PROGRESS")=tpTbl3Unpassed
		rs("TASK_EVAL")=eval_text
	End If
	rs.Update()
	CloseRs rs
	CloseConn conn
Case 4	'  审核答辩材料
	filetypename="答辩审批材料"
	If ispass Then
		rs("TASK_PROGRESS")=tpTbl4Passed
	Else
		rs("TASK_PROGRESS")=tpTbl4Unpassed
	End If
	rs("TASK_EVAL")=eval_text
	rs.Update()
	CloseRs rs
	CloseConn conn
Case 5	'  同意/不同意送检送审操作
	filetypename="送检论文和送审论文"
	author=Request.Form("author")
	stuno=Request.Form("stuno")
	tutorinfo=Request.Form("tutorinfo")
	speciality=Request.Form("speciality")
	faculty=Request.Form("faculty")
	subject=Request.Form("subject")
	If Not ispass And (Len(author)=0 Or Len(stuno)=0 Or Len(tutorinfo)=0 Or Len(speciality)=0 Or Len(faculty)=0 _
	Or Len(subject)=0) Then
		bError=True
		errdesc="缺少必要的字段信息！"
	ElseIf review_status>=rsNotAgreeDetect Then
		bError=True
		errdesc="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		CloseRs rs
  	CloseConn conn
  	Response.End
	End If
	' 更新记录
	If ispass Then
		rs("DETECT_APP_EVAL")="论文已检查，同意检测。"
		rs("REVIEW_APP_EVAL")=eval_text
		rs("SUBMIT_REVIEW_TIME")=Now
		rs("REVIEW_STATUS")=rsAgreeDetect
	Else
		rs("DETECT_APP_EVAL")=eval_text
		rs("REVIEW_APP_EVAL")=Null
		rs("REVIEW_STATUS")=rsNotAgreeDetect
	End If
	rs.Update()
	CloseRs rs
	CloseConn conn
Case 6	'  同意/不同意送审操作
	filetypename="送审论文"
	author=Request.Form("author")
	stuno=Request.Form("stuno")
	tutorinfo=Request.Form("tutorinfo")
	speciality=Request.Form("speciality")
	faculty=Request.Form("faculty")
	subject=Request.Form("subject")
	reproduct_ratio=Request.Form("reproduct_ratio")
	If Len(author)=0 Or Len(stuno)=0 Or Len(tutorinfo)=0 Or Len(speciality)=0 Or Len(faculty)=0 _
	Or Len(subject)=0 Or ispass And Len(reproduct_ratio)=0 Then
		bError=True
		errdesc="缺少必要的字段信息！"
	ElseIf review_status>=rsNotAgreeReview Then
		bError=True
		errdesc="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		CloseRs rs
  	CloseConn conn
  	Response.End
	End If
	If ispass Then
		' 生成送审申请表
		Dim rag,review_time
		review_time=Now
		Randomize
		filename=toDateTime(review_time,1)&Int(Timer)&Int(Rnd()*999)&".doc"
		filepath=Server.MapPath("export")&"\"&filename
		Set rag=New ReviewAppGen
		rag.Author=author
		rag.StuNo=stuno
		rag.TutorInfo=tutorinfo
		rag.Spec=speciality
		rag.Date=toDateTime(review_time,1)
		rag.Subject=subject
		rag.EvalText=eval_text
		rag.ReproductRatio=reproduct_ratio
		bError=rag.generateApp(filepath)=0
		Set rag=Nothing
		rs("REVIEW_APP")=filename
		rs("SUBMIT_REVIEW_TIME")=review_time
		rs("REVIEW_STATUS")=rsAgreeReview
	Else
		rs("REVIEW_STATUS")=rsNotAgreeReview
	End If
	' 更新记录
	rs("REVIEW_APP_EVAL")=eval_text
	rs.Update()
	CloseRs rs
	CloseConn conn
Case 7	' 评阅书审阅确认操作
	If review_status>=rsReviewEval Then
		bError=True
		errdesc="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		CloseRs rs
  	CloseConn conn
  	Response.End
	End If
	' 更新记录
	rs("TUTOR_REVIEW_EVAL_TIME")=Now
	rs("REVIEW_STATUS")=rsReviewEval
	rs.Update()
	CloseRs rs
	CloseConn conn
Case 8	'  提交答辩论文审核意见操作
	filetypename="答辩论文"
	If review_status>=rsModifyUnpassed Then
		bError=True
		errdesc="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		CloseRs rs
  	CloseConn conn
  	Response.End
	End If
	' 更新记录
	If ispass Then
		'eval_text="已审阅，同意答辩"
		rs("REVIEW_STATUS")=rsModifyPassed
	Else
		'eval_text="不同意答辩，请继续修改论文"
		rs("REVIEW_STATUS")=rsModifyUnpassed
	End If
	rs("TUTOR_MODIFY_EVAL")=eval_text
	rs("TUTOR_MODIFY_EVAL_TIME")=Now
	rs.Update()
	CloseRs rs
	CloseConn conn
End Select
If opr=7 Then
	' 向学生发送评阅意见确认通知邮件
	sendEmailToStudent thesisID,"",True,""
Else
	' 向学生发送审核结果通知邮件
	sendEmailToStudent thesisID,filetypename,ispass,eval_text
End If
updateActiveTime Session("Tid")

%><form id="ret" action="thesisDetail.asp?tid=<%=thesisID%>" method="post">
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID2" value="<%=spec_id%>" />
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
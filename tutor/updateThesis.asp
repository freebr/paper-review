<%Response.Expires=-1%>
<!--#include file="../inc/global.inc"-->
<!--#include file="appgen.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Tid")) Then Response.Redirect("../error.asp?timeout")
step=Request.QueryString("step")
thesisID=Request.QueryString("tid")
opr=Request.Form("opr")
submittype=Request.Form("submittype")
is_pass=submittype="pass"
eval_text=Request.Form("eval_text")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
spec_id=Request.Form("In_SPECIALITY_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
If Len(thesisID)=0 Or Not IsNumeric(thesisID) Then
	bError=True
	errdesc="参数无效。"
ElseIf Not isMatched("[1-8]",opr,True) Then
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
	CloseRs rs
	CloseConn conn
	showErrorPage errdesc, "提示"
End If

Dim conn,sql,ret,rs,count
Connect conn
sql="SELECT * FROM ViewDissertations WHERE ID=?"
Set ret=ExecQuery(conn,sql,CmdParam("ID",adInteger,4,thesisID))
Set rs=ret("rs")
If rs.EOF Then
  	CloseRs rs
  	CloseConn conn
	showErrorPage "数据库没有该论文记录！", "提示"
End If

Dim section_access_info
Set section_access_info=getSectionAccessibilityInfo(rs("ActivityId"),rs("TEACHTYPE_ID"),sectionAudit)
If Not section_access_info("accessible") Then
	CloseRs rs
  	CloseConn conn
	showErrorPage section_access_info("tip"), "提示"
End If

CloseRs rs
sql="SELECT * FROM Dissertations WHERE ID="&thesisID
GetRecordSet conn,rs,sql,count
review_status=rs("REVIEW_STATUS")
Select Case opr
Case 1	'	 审核开题报告表
	file_type_name="开题报告表/开题论文"
	If is_pass Then
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
	file_type_name="中期检查表/中期论文"
	If is_pass Then
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
	file_type_name="预答辩申请表/预答辩论文"
	If is_pass Then
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
	file_type_name="答辩审批材料"
	If is_pass Then
		rs("TASK_PROGRESS")=tpTbl4Passed
	Else
		rs("TASK_PROGRESS")=tpTbl4Unpassed
	End If
	rs("TASK_EVAL")=eval_text
	rs.Update()
	CloseRs rs
	CloseConn conn
Case 5	'  同意/不同意送检送审操作
	file_type_name="送检论文和送审论文"
	author=Request.Form("author")
	stuno=Request.Form("stuno")
	tutorinfo=Request.Form("tutorinfo")
	speciality=Request.Form("speciality")
	faculty=Request.Form("faculty")
	subject=Request.Form("subject")
	If Not is_pass And (Len(author)=0 Or Len(stuno)=0 Or Len(tutorinfo)=0 Or Len(speciality)=0 Or Len(faculty)=0 _
	Or Len(subject)=0) Then
		bError=True
		errdesc="缺少必要的字段信息！"
	ElseIf review_status>=rsNotAgreeDetect Then
		bError=True
		errdesc="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		CloseRs rs
  		CloseConn conn
		showErrorPage errdesc, "提示"
	End If
	' 更新记录
	If is_pass Then
		sql="SELECT dbo.getDetectResultCount("&thesisID&")"
		GetRecordSet conn,rsDetect,sql,count
		detect_count=rsDetect(0).Value
		If detect_count>=1 Then
			rs("DETECT_APP_EVAL")="该生已对论文进行修改，并已经导师检查，同意二次检测。"
		Else
			rs("DETECT_APP_EVAL")="论文已检查，同意检测。"
		End If
		rs("REVIEW_APP_EVAL")=eval_text
		rs("SUBMIT_REVIEW_TIME")=Now
		rs("REVIEW_STATUS")=rsAgreeDetect
		CloseRs rsDetect
	Else
		rs("DETECT_APP_EVAL")=eval_text
		rs("REVIEW_APP_EVAL")=Null
		rs("SUBMIT_REVIEW_TIME")=Now
		rs("REVIEW_STATUS")=rsNotAgreeDetect
	End If
	rs.Update()
	CloseRs rs
	CloseConn conn
Case 6	'  同意/不同意送审操作
	file_type_name="送审论文"
	author=Request.Form("author")
	stuno=Request.Form("stuno")
	tutorinfo=Request.Form("tutorinfo")
	speciality=Request.Form("speciality")
	faculty=Request.Form("faculty")
	subject=Request.Form("subject")
	reproduct_ratio=Request.Form("reproduct_ratio")
	If Len(author)=0 Or Len(stuno)=0 Or Len(tutorinfo)=0 Or Len(speciality)=0 Or Len(faculty)=0 _
	Or Len(subject)=0 Or is_pass And Len(reproduct_ratio)=0 Then
		bError=True
		errdesc="缺少必要的字段信息！"
	ElseIf review_status>=rsNotAgreeReview Then
		bError=True
		errdesc="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		CloseRs rs
  		CloseConn conn
		showErrorPage errdesc, "提示"
	End If
	If is_pass Then
		' 生成送审申请表
		Dim rag,review_time
		review_time=Now
		Randomize()
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
		CloseRs rs
  		CloseConn conn
		showErrorPage errdesc, "提示"
	End If
	' 更新记录
	rs("TUTOR_REVIEW_EVAL_TIME")=Now
	rs("REVIEW_STATUS")=rsReviewEval
	rs.Update()
	CloseRs rs
	CloseConn conn
Case 8	'  提交答辩论文审核意见操作
	file_type_name="答辩论文"
	If review_status>=rsModifyUnpassed Then
		bError=True
		errdesc="本论文当前状态下不能执行此操作！"
	End If
	If bError Then
		CloseRs rs
  		CloseConn conn
		showErrorPage errdesc, "提示"
	End If
	' 更新记录
	If is_pass Then
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
	sendEmailToStudent thesisID,file_type_name,is_pass,eval_text
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
<%Response.Charset="utf-8"
Response.Expires=-1%>
<!--#include file="../inc/upload_5xsoft.inc"-->
<!--#include file="appgen.inc"-->
<!--#include file="evalappend.inc"-->
<!--#include file="../inc/db.asp"-->
<!--#include file="../inc/mail.asp"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("user")) Then Response.Redirect("../error.asp?timeout")
Dim Upload
Set Upload=New upload_5xsoft
curstep=Request.QueryString("step")
thesisID=Request.QueryString("tid")
new_subject_ch=Upload.Form("new_subject_ch")
new_subject_en=Upload.Form("new_subject_en")
new_keywords_ch=Upload.Form("new_keywords_ch")
new_keywords_en=Upload.Form("new_keywords_en")
new_period_id=Upload.Form("new_period_id")
new_reviewfilestat=Upload.Form("new_reviewfilestat")
new_task_progress=Upload.Form("new_task_progress")
new_review_status=Upload.Form("new_review_status")
new_reproduct_ratio=Upload.Form("reproduct_ratio")
new_defence_result=Upload.Form("defenceresult")
new_grant_degree=Upload.Form("grantdegree")
opr=Upload.Form("opr")
submittype=Upload.Form("submittype")
ispass=submittype="pass"
eval_text=Upload.Form("eval_text")
Set detect_report=Upload.File("detectreport")
period_id=Upload.Form("In_PERIOD_ID2")
teachtype_id=Upload.Form("In_TEACHTYPE_ID2")
class_id=Upload.Form("In_CLASS_ID2")
enter_year=Upload.Form("In_ENTER_YEAR2")
query_task_progress=Upload.Form("In_TASK_PROGRESS2")
query_review_status=Upload.Form("In_REVIEW_STATUS2")
finalFilter=Upload.Form("finalFilter2")
pageSize=Upload.Form("pageSize2")
pageNo=Upload.Form("pageNo2")
If Len(thesisID)=0 Or Not IsNumeric(thesisID) Or Not IsNumeric(opr) Then
	bError=True
	errdesc="参数无效。"
ElseIf submittype<>vbNullString And Not isMatched("[0-8]",opr) Then
	bError=True
	errdesc="操作无效。"
ElseIf submittype="unpass" And opr<=5 Or submittype<>vbNullString And (opr=4 Or opr=6) Then
	If Len(eval_text)=0 Then
		bError=True
		errdesc="请填写意见（200-2000字）！"
	ElseIf Len(eval_text)>2000 Then
		bError=True
		errdesc="意见字数超出限制（2000字）！"
	End If
ElseIf new_reproduct_ratio<>vbNullString And Not IsNumeric(new_reproduct_ratio) Then
	bError=True
	errdesc="复制比输入无效，请输入 0-100 间的数字！"
ElseIf Not isMatched("[0-4]",new_defence_result) Then
	bError=True
	errdesc="答辩成绩输入无效！"
ElseIf Not isMatched("[0-2]",new_grant_degree) Then
	bError=True
	errdesc="“是否同意授予学位”设置无效！"
ElseIf detect_report.FileName<>vbNullString And new_reproduct_ratio=vbNullString Then
	bError=True
	errdesc="请填写复制比！"
End If
If bError Then
%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	CloseRs rs
	CloseConn conn
	Response.End
End If

Dim conn,rs,sql,result
Connect conn
sql="SELECT * FROM TEST_THESIS_REVIEW_INFO WHERE ID="&thesisID
GetRecordSet conn,rs,sql,result
If rs.EOF Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">数据库没有该论文记录！</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
  CloseRs rs
  CloseConn conn
	Response.End
End If

If submittype=vbNullString Then
	opr=0
End If
review_status=rs("REVIEW_STATUS")
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
Case 2	'  审核中期检查表
	filetypename="中期检查表/中期论文"
	If ispass Then
		rs("TASK_PROGRESS")=tpTbl2Passed
		rs("TASK_EVAL")=Null
	Else
		rs("TASK_PROGRESS")=tpTbl2Unpassed
		rs("TASK_EVAL")=eval_text
	End If
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
Case 4	'  审核答辩材料
	filetypename="答辩审批材料"
	If ispass Then
'		sql="SELECT TEACHTYPE_ID,TUTOR_NAME FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE ID="&thesisID
'		GetRecordSetNoLock conn,rs2,sql,result
'		teachtype_id=rs2(0)
'		tutor_name=rs2(1)
'		CloseRs rs2
'		' 追加导师意见
'		Dim evalap,append_time
'		append_time=Now
'		'Set fso=Server.CreateObject("Scripting.FileSystemObject")
'		Set evalap=New EvalAppend
'		evalap.TutorName=tutor_name
'		evalap.Date=append_time
'		evalap.EvalText=eval_text
'		Select Case teachtype_id
'		Case 5:filetype=1
'		Case 6,9:filetype=2
'		End Select
'		uploadpath=Server.MapPath("/ThesisReview/student/upload")&"/"
'		filepath=uploadpath&rs("TABLE_FILE4")
'		tempfilename=FormatDateTime(append_time,1)&Int(Timer)&Int(Rnd()*999)&".doc"
'		tempfile=uploadpath&tempfilename
'		'fso.CopyFile filepath,tempfile
'		bError=evalap.appendEval(tempfile,filepath,filetype)=0
'		'fso.DeleteFile filepath
'		Set evalap=Nothing
'		'Set fso=Nothing
'		rs("TABLE_FILE4")=tempfilename
		rs("TASK_PROGRESS")=tpTbl4Passed
	Else
		rs("TASK_PROGRESS")=tpTbl4Unpassed
	End If
	rs("TASK_EVAL")=eval_text
Case 5	'  同意/不同意送检操作
	filetypename="送检论文"
	author=Upload.Form("author")
	stuno=Upload.Form("stuno")
	tutorinfo=Upload.Form("tutorinfo")
	speciality=Upload.Form("speciality")
	faculty=Upload.Form("faculty")
	subject=Upload.Form("subject")
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
	If ispass Then
		sql="SELECT STU_NAME,STU_NO,TUTOR_ID,TUTOR_NAME,SPECIALITY_NAME,TEACHTYPE_NAME,THESIS_SUBJECT,MOBILE,EMAIL,TUTOR_MOBILE,TUTOR_EMAIL FROM VIEW_TEST_THESIS_REVIEW_INFO "&_
				"WHERE ID="&thesisID
		GetRecordSetNoLock conn,rs2,sql,result
		If result=1 Then
			author=rs2("STU_NAME")
			stuno=rs2("STU_NO")
			tutorname=rs2("TUTOR_NAME")
			tutorinfo=tutorname&" "&getProDutyNameOf(rs2("TUTOR_ID"))
			speciality=rs2("SPECIALITY_NAME")
			degreename=rs2("TEACHTYPE_NAME")
			subject=rs2("THESIS_SUBJECT")
			mobile=rs2("MOBILE")
			email=rs2("EMAIL")
			tutormobile=rs2("TUTOR_MOBILE")
			tutormail=rs2("TUTOR_EMAIL")
		End If
		CloseRs rs2
		eval_text="论文已检查，同意检测。"
		' 更新记录
		rs("REVIEW_STATUS")=rsAgreeDetect
	Else
		rs("REVIEW_STATUS")=rsNotAgreeDetect
	End If
	rs("DETECT_APP_EVAL")=eval_text
Case 6	'  同意/不同意送审操作
	filetypename="送审论文"
	author=Upload.Form("author")
	stuno=Upload.Form("stuno")
	tutorinfo=Upload.Form("tutorinfo")
	speciality=Upload.Form("speciality")
	faculty=Upload.Form("faculty")
	subject=Upload.Form("new_subject")
	reproduct_ratio=Upload.Form("reproduct_ratio")
	If Len(author)=0 Or Len(stuno)=0 Or Len(tutorinfo)=0 Or Len(speciality)=0 Or Len(faculty)=0 _
	Or Len(subject)=0 Or ispass And Len(reproduct_ratio)=0 Then
		'bError=True
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
		filename=FormatDateTime(review_time,1)&Int(Timer)&Int(Rnd()*999)&".docx"
		filepath=Server.MapPath("/ThesisReview/teacher/export")&"\"&filename
		Set rag=New ReviewAppGen
		rag.Author=author
		rag.StuNo=stuno
		rag.TutorInfo=tutorinfo
		rag.Spec=speciality
		rag.Date=FormatDateTime(review_time,1)
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
Case 8	'  提交论文修改意见操作
	filetypename="修改后论文"
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
		eval_text="已审阅，同意答辩"
		rs("REVIEW_STATUS")=rsModifyPassed
	Else
		eval_text="不同意答辩，请继续修改论文"
		rs("REVIEW_STATUS")=rsModifyUnpassed
	End If
	rs("TUTOR_MODIFY_EVAL")=eval_text
	rs("TUTOR_MODIFY_EVAL_TIME")=Now
End Select
If submittype=vbNullString Then
	' 更新表单信息
	If Len(new_reviewfilestat) Then
		rs("REVIEW_FILE_STATUS")=new_reviewfilestat
	End If
	If Len(new_task_progress) Then
		rs("TASK_PROGRESS")=new_task_progress
	End If
	If Len(new_review_status) Then
		rs("REVIEW_STATUS")=new_review_status
	End If
	If detect_report.FileName<>vbNullString Then
		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		reportDir=getDateTimeId(Now)
		strUploadPath=Server.MapPath("upload/report/"&reportDir)
		' 检查上传目录是否存在
		If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)
		fileExt=LCase(detect_report.FileExt)
		' 生成日期格式文件名
		fileid=FormatDateTime(Now(),1)&Int(Timer)
		strDestFile=fileid&"."&fileExt
		strDestPath=strUploadPath&"\"&strDestFile
		' 保存
		detect_report.SaveAs strDestPath
		rs("DETECT_REPORT")=reportDir&"/"&strDestFile
		Set fso=Nothing
	End If
	reproduct_ratio=rs("REPRODUCTION_RATIO")
	If Not IsNull(reproduct_ratio) And new_reproduct_ratio=vbNullString And review_status>=rsAgreeDetect Then
		rs("REVIEW_STATUS")=rsAgreeDetect
	ElseIf new_reproduct_ratio<>vbNullString And new_review_status<rsDetected Then
		rs("REVIEW_STATUS")=rsDetected
	End If
	If Len(new_reproduct_ratio)=0 Then
		rs("REPRODUCTION_RATIO")=Null
	Else
		rs("REPRODUCTION_RATIO")=new_reproduct_ratio
	End If
	If Len(new_period_id)=0 Then
		rs("PERIOD_ID")=Null
	Else
		rs("PERIOD_ID")=new_period_id
	End If
	If Len(new_defence_result)<>0 Then
		sql="UPDATE TEST_THESIS_DEFENCE_INFO SET DEFENCE_RESULT="&new_defence_result&" WHERE THESIS_ID="&thesisID
		conn.Execute sql
	End If
	If Len(new_grant_degree)<>0 Then
		rs("GRANT_DEGREE")=Array(Null,True,False)(new_grant_degree)
	End If
	If Len(new_subject_ch)=0 Then
		rs("THESIS_SUBJECT")=Null
	Else
		rs("THESIS_SUBJECT")=new_subject_ch
	End If
	If Len(new_subject_en)=0 Then
		rs("THESIS_SUBJECT_EN")=Null
	Else
		rs("THESIS_SUBJECT_EN")=new_subject_en
	End If
	If Len(new_keywords_ch)=0 Then
		rs("KEYWORDS")=Null
	Else
		rs("KEYWORDS")=new_keywords_ch
	End If
	If Len(new_keywords_en)=0 Then
		rs("KEYWORDS_EN")=Null
	Else
		rs("KEYWORDS_EN")=new_keywords_en
	End If
End If
rs.Update()
CloseRs rs
CloseConn conn
If opr=7 Then
	' 向学生发送修改论文通知邮件
	sendEmailToStudent thesisID,"",True,""
ElseIf opr<>0 Then
	' 向学生发送审核结果通知邮件
	sendEmailToStudent thesisID,filetypename,ispass,eval_text
End If
%><form id="ret" action="thesisDetail.asp?tid=<%=thesisID%>" method="post">
<input type="hidden" name="In_PERIOD_ID2" value="<%=period_id%>">
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
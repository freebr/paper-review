<%Response.Expires=-1%>
<!--#include file="reviewgen.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")
dissertation_id=Request.QueryString("tid")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
spec_id=Request.Form("In_SPECIALITY_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
If Len(dissertation_id)=0 Or Not IsNumeric(dissertation_id) Then
	bError=True
	errdesc="参数无效。"
ElseIf Not checkIfProfileFilledIn() Then
	bError=True
	errdesc="您尚未完善个人信息，<a href=""profile.asp"">请点击这里编辑。</a>"
End If
If bError Then
	showErrorPage errdesc, "提示"
End If
master_level=Request.Form("master_level")
correlation_level=Request.Form("correlation_level")
review_result=Request.Form("review_result")
review_level=Request.Form("review_level")
eval_text=Request.Form("eval_text")

Dim conn,sql,ret,rs,count
Connect conn
sql="SELECT * FROM ViewThesisInfo WHERE ID=? AND ? IN (REVIEWER1,REVIEWER2)"
Set ret=ExecQuery(conn,sql,_
	CmdParam("ID",adInteger,4,dissertation_id),_
	CmdParam("TId",adInteger,4,Session("TId")))
Set rs=ret("rs")
count=ret("count")
If Len(master_level)=0 Then
	bError=True
	errdesc="请选择您对论文涉及内容的熟悉程度！"
ElseIf Len(correlation_level)=0 Then
	bError=True
	errdesc="请选择学位论文内容与申请学位专业（领域）的相关性！"
ElseIf Len(review_result)=0 Or review_result="0" Then
	bError=True
	errdesc="请就是否同意举行论文答辩选择相应选项！"
ElseIf InStr("1234",review_result)=0 Then
	bError=True
	errdesc="“是否同意举行论文答辩”选项无效！"
ElseIf Len(eval_text)=0 Then
	bError=True
	errdesc="请填写评语（200-2000字）！"
ElseIf Len(eval_text)>2000 Then
	bError=True
	errdesc="评语字数超出限制（2000字）！"
ElseIf count=0 Then
	bError=True
	errdesc="数据库没有该论文记录，或您未受邀评阅该论文！"
ElseIf rs("REVIEW_STATUS")<>rsMatchExpert And rs("REVIEW_STATUS")<>rsReviewed Then
	bError=True
	errdesc="本论文当前状态下不允许评阅！"
Else
	For i=1 To Request.Form("scores").Count
		n=Request.Form("scores")(i)
		If Len(n)=0 Or Not IsNumeric(n) Then
			bError=True
		ElseIf n<0 Or n>100 Then
			bError=True
		ElseIf InStr(n,".") Then
			bError=True
		End If
		If bError Then
			errdesc="第&nbsp;"&i&"&nbsp;项得分值无效，请输入0-100之间的整数！"
			Exit For
		End If
	Next
	If Not bError Then
		If Len(review_level)=0 Then
			bError=True
			errdesc="缺少总体评价参数！"
		End If
	End If
End If
If bError Then
	CloseRs rs
	CloseConn conn
	showErrorPage errdesc, "提示"
End If

Dim section_access_info
Set section_access_info=getSectionAccessibilityInfo(rs("ActivityId"),rs("TEACHTYPE_ID"),sectionReview)
If Not section_access_info("accessible") Then
	CloseRs rs
  	CloseConn conn
	showErrorPage section_access_info("tip"), "提示"
End If

Dim display_status,reviewer_num
Dim arr_master_level(1)
Dim arr_review_level(1)
Dim arr_review_result(2)
Dim arr_review_file(1)
Dim arr_review_time(1)

display_status=rs("REVIEW_FILE_STATUS")
If rs("REVIEWER1")=Session("TId") Then
	reviewer_num=0
Else
	reviewer_num=1
End If
If Not IsNull(rs("REVIEW_RESULT")) Then
	arr=Split(rs("REVIEW_RESULT"),",")
	For i=0 To UBound(arr)
		arr_review_result(i)=Int(arr(i))
	Next
End If
If IsNull(rs("REVIEWER_EVAL_TIME")) Then
	For i=0 To 1
		arr_master_level(i)=0
		arr_review_level(i)=0
	Next
Else
	arr2=Split(rs("REVIEWER_MASTER_LEVEL"),",")
	arr3=Split(rs("REVIEW_FILE"),",")
	arr4=Split(rs("REVIEWER_EVAL_TIME"),",")
	arr5=Split(rs("REVIEW_LEVEL"),",")
	For i=0 To 1
		arr_master_level(i)=Int(arr2(i))
		arr_review_file(i)=arr3(i)
		arr_review_time(i)=arr4(i)
		arr_review_level(i)=Int(arr5(i))
	Next
End If
teachtype_id=rs("TEACHTYPE_ID")
review_type=rs("REVIEW_TYPE")
submit_review_date=toDateTime(rs("SUBMIT_REVIEW_TIME"),1)
author=rs("STU_NAME")
tutorinfo=rs("TUTOR_NAME")&" "&getProDutyNameOf(rs("TUTOR_ID"))
subject=rs("THESIS_SUBJECT")
speciality=rs("SPECIALITY_NAME")
researchway=rs("RESEARCHWAY_NAME")
scores=Request.Form("scores")
CloseRs rs
sql="SELECT * FROM Experts WHERE TEACHER_ID="&Session("TId")
GetRecordSetNoLock conn,rs,sql,count
expert_name=rs("EXPERT_NAME")
expert_pro_duty=rs("PRO_DUTY_NAME")
expert_expertise=rs("EXPERTISE")
expert_workplace=rs("WORKPLACE")
expert_address=rs("ADDRESS")
expert_mailcode=rs("MAILCODE")
expert_telephone=rs("TELEPHONE")
expert_mobile=rs("MOBILE")
CloseRs rs

sql="SELECT REVIEW_FILE FROM ReviewTypes WHERE ID="&review_type
GetRecordSetNoLock conn,rs,sql,count
If rs.EOF Then
	CloseRs rs
	CloseConn conn
	showErrorPage "评阅书模板丢失，无法完成评阅操作，请联系系统管理员。", "提示"
End If
' 生成评阅书
Dim rg,review_time,template_file,reviewfile_type,filepath,filename,full_filename
template_file=Server.MapPath("../admin/upload/review/"&rs("REVIEW_FILE"))
CloseRs rs
review_time=Now()
If teachtype_id=5 Or teachtype_id=6 Then
	reviewfile_type=2
Else
	reviewfile_type=1
End If
Randomize()
filename=toDateTime(review_time,1)&Int(Timer)&Int(Rnd()*999)
full_filename=filename&".pdf"
filepath=Server.MapPath("export")&"\"&full_filename
filepath2=Server.MapPath("export")&"\"&filename&"_nostu.pdf"
filepath3=Server.MapPath("export")&"\"&filename&"_noexp.pdf"

Set rg=New ReviewGen
rg.Author=author
rg.TutorInfo=tutorinfo
rg.Subject=subject
rg.ResearchWay=researchway
rg.Date=submit_review_date
rg.ExpertName=expert_name
rg.ExpertProDuty=expert_pro_duty
rg.ExpertExpertise=expert_expertise
rg.ExpertWorkplace=expert_workplace
rg.ExpertAddress=expert_address
rg.ExpertMailcode=expert_mailcode
rg.ExpertTel1=expert_telephone
rg.ExpertTel2=expert_mobile
rg.ExpertMasterLevel=master_level
rg.EvalText=eval_text
rg.CorrelationLevel=correlation_level
rg.ReviewResult=review_result
rg.ReviewLevel=review_level
rg.ThesisType=review_type
If reviewfile_type=2 Then	' ME/MBA评阅书，计算评价指标总分
	rg.Spec=speciality
	rg.Scores=scores
	Dim arrScorePartPower,arrScores,arrScorePower
	Dim scoreParts,partScore,totalScore
	loadReviewScoringInfo review_type,tmp,code_power1,code_power2
	code_power1=Replace(code_power1,"[","Array(")
	code_power1=Replace(code_power1,"]",")")
	code_power2=Replace(code_power2,"[","Array(")
	code_power2=Replace(code_power2,"]",")")
	arrScorePartPower=Eval(code_power1)
	arrScorePower=Eval(code_power2)
	arrScores=Split(scores,",")
	totalScore=0
	k=0
	For i=0 To UBound(arrScorePartPower)
		partScore=0
		For j=0 To UBound(arrScorePower(i))
			arrScores(k)=Int(arrScores(k))
			partScore=partScore+arrScores(k)*arrScorePower(i)(j)
			k=k+1
		Next
		If i>0 Then scoreParts=scoreParts&","
		partScore=partScore*arrScorePartPower(i)
		scoreParts=scoreParts&partScore
		totalScore=totalScore+partScore
	Next
	rg.ScoreParts=scoreParts
	rg.TotalScore=Round(totalScore)
	score_data=Join(arrScores,",")
Else
	score_data=Null
End If
bError=rg.generateReview(filepath,filepath2,filepath3,template_file,reviewfile_type)=0
Set rg=Nothing

arr_master_level(reviewer_num)=master_level
arr_review_level(reviewer_num)=review_level
arr_review_result(reviewer_num)=review_result
arr_review_file(reviewer_num)=full_filename
arr_review_time(reviewer_num)=review_time
' 确定处理意见
code=arr_review_result(0)&arr_review_result(1)
Select Case code
Case "11":finalresult="1"
Case "12","21":finalresult="2"
Case "22":finalresult="2"
Case "13","31":finalresult="3"
Case "23","32":finalresult="3"
Case "33":finalresult="5"
Case "14","41":finalresult="4"
Case "24","42":finalresult="4"
Case "34","43":finalresult="5"
Case "44":finalresult="5"
Case Else:finalresult="6"
End Select
arr_review_result(2)=finalresult

' 插入评阅记录
sql="EXEC spAddReviewRecord ?,?,?,?,?,?,?,?,?,?,?,NULL"
ExecNonQuery conn,sql,_
	CmdParam("dissertation_id",adInteger,4,dissertation_id),_
	CmdParam("reviewer_id",adInteger,4,Session("TId")),_
	CmdParam("reviewer_master_level",adInteger,4,master_level),_
	CmdParam("score_data",adVarWChar,500,score_data),_
	CmdParam("comment",adLongVarWChar,5000,eval_text),_
	CmdParam("correlation_level",adInteger,4,correlation_level),_
	CmdParam("overall_rating",adInteger,4,review_level),_
	CmdParam("defence_opinion",adInteger,4,review_result),_
	CmdParam("review_time",adDate,4,review_time),_
	CmdParam("review_file",adVarWChar,50,filename),_
	CmdParam("display_status",adInteger,4,display_status)

' 更新记录
sql="SELECT * FROM Dissertations WHERE ID="&dissertation_id
GetRecordSet conn,rs,sql,count
rs("REVIEW_RESULT")=ArrayJoin(arr_review_result,",")
rs("REVIEW_LEVEL")=ArrayJoin(arr_review_level,",")
rs("REVIEWER_MASTER_LEVEL")=ArrayJoin(arr_master_level,",")
rs("REVIEWER_EVAL"&(reviewer_num+1))=eval_text
rs("REVIEW_FILE")=ArrayJoin(arr_review_file,",")
rs("REVIEWER_EVAL_TIME")=ArrayJoin(arr_review_time,",")
If finalresult<>"6" Then
	rs("REVIEW_STATUS")=rsReviewed
End If
rs.Update()
CloseRs rs
CloseConn conn

updateActiveTime Session("TId")

logtxt=Format("专家[{0}]提交评阅意见，论文：《{1}》，作者：{2}，评阅书：{3}。",_
	expert_name,subject,author,full_filename)
writeLog logtxt
%><form id="ret" action="thesisDetail.asp?tid=<%=dissertation_id%>" method="post">
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID2" value="<%=spec_id%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" />
<input type="hidden" name="finishReview" value="1" /></form>
<script type="text/javascript">
	alert("提交成功，感谢您参与本论文评阅！");
	document.all.ret.submit();
</script>
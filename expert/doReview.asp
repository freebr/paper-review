<%Response.Charset="utf-8"
Response.Expires=-1%>
<!--#include file="reviewgen.inc"-->
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")
thesisID=Request.QueryString("tid")
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
ElseIf Not checkIfProfileFilledIn() Then
	bError=True
	errdesc="您尚未完善个人信息，<a href=""profile.asp"">请点击这里编辑。</a>"
End If
If bError Then
%><html><head><link href="../css/tutor.css" rel="stylesheet" type="text/css" /></head>
<body class="exp"><center><div class="content"><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></div></center></body></html><%
	Response.End()
End If
expert_master_level=Request.Form("masterlevel")
correlation_type=Request.Form("correlationtype")
reviewresult=Request.Form("reviewresult")
reviewlevel=Request.Form("reviewlevel")
eval_text=Request.Form("eval_text")
Session("Debug")=1
Connect conn
sql="SELECT * FROM ViewThesisInfo WHERE ID="&thesisID&" AND "&Session("Tid")&" IN (REVIEWER1,REVIEWER2)"
GetRecordSetNoLock conn,rs,sql,result
If nSystemStatus<>2 Then
	bError=True
	errdesc="评阅系统的开放时间为"&toDateTime(startdate,1)&"至"&toDateTime(enddate,1)&"，当前不在开放时间内，不能评阅论文。"
ElseIf Len(expert_master_level)=0 Then
	bError=True
	errdesc="请选择您对论文涉及内容的熟悉程度！"
ElseIf Len(correlation_type)=0 Then
	bError=True
	errdesc="请选择学位论文内容与申请学位专业（领域）的相关性！"
ElseIf Len(reviewresult)=0 Or reviewresult="0" Then
	bError=True
	errdesc="请就是否同意举行论文答辩选择相应选项！"
ElseIf InStr("1234",reviewresult)=0 Then
	bError=True
	errdesc="“是否同意举行论文答辩”选项无效！"
ElseIf Len(eval_text)=0 Then
	bError=True
	errdesc="请填写评语（200-2000字）！"
ElseIf Len(eval_text)>2000 Then
	bError=True
	errdesc="评语字数超出限制（2000字）！"
ElseIf result=0 Then
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
		If Len(reviewlevel)=0 Then
			bError=True
			errdesc="缺少总体评价参数！"
		End If
	End If
End If
If bError Then
%><html><head><link href="../css/tutor.css" rel="stylesheet" type="text/css" /></head>
<body class="exp"><center><div class="content"><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></div></center></body></html><%
  CloseRs rs
  CloseConn conn
	Response.End()
End If

Dim reviewer_type
Dim review_result(2),reviewer_master_level(1),review_file(1),review_time(1),review_level(1)
If rs("REVIEWER1")=Session("Tid") Then
	reviewer_type=0
Else
	reviewer_type=1
End If
If Not IsNull(rs("REVIEW_RESULT")) Then
	arr=Split(rs("REVIEW_RESULT"),",")
	For i=0 To UBound(arr)
		review_result(i)=Int(arr(i))
	Next
End If
If IsNull(rs("REVIEWER_EVAL_TIME")) Then
	For i=0 To 1
		reviewer_master_level(i)=0
		review_level(i)=0
	Next
Else
	arr2=Split(rs("REVIEWER_MASTER_LEVEL"),",")
	arr3=Split(rs("REVIEW_FILE"),",")
	arr4=Split(rs("REVIEWER_EVAL_TIME"),",")
	arr5=Split(rs("REVIEW_LEVEL"),",")
	For i=0 To 1
		reviewer_master_level(i)=Int(arr2(i))
		review_file(i)=arr3(i)
		review_time(i)=arr4(i)
		review_level(i)=Int(arr5(i))
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
sql="SELECT * FROM TEST_THESIS_REVIEW_EXPERT_INFO WHERE TEACHER_ID="&Session("Tid")
GetRecordSetNoLock conn,rs,sql,result
expert_name=rs("EXPERT_NAME")
expert_pro_duty=rs("PRO_DUTY_NAME")
expert_expertise=rs("EXPERTISE")
expert_workplace=rs("WORKPLACE")
expert_address=rs("ADDRESS")
expert_mailcode=rs("MAILCODE")
expert_telephone=rs("TELEPHONE")
expert_mobile=rs("MOBILE")
CloseRs rs

sql="SELECT REVIEW_FILE FROM CODE_REVIEW_TYPE WHERE ID="&review_type
GetRecordSetNoLock conn,rs,sql,result
If rs.EOF Then
%><html><head><link href="../css/tutor.css" rel="stylesheet" type="text/css" /></head>
<body class="exp"><center><div class="content"><body bgcolor="ghostwhite"><center><font color=red size="4">操作不成功，找不到所需的评阅书模板文件，请联系系统管理员。</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></div></center></body></html><%
  CloseRs rs
  CloseConn conn
	Response.End()
End If
' 生成评阅书
Dim rg,reviewtime,templatefile,reviewfile_type,filepath,filename,fullfilename
templatefile=Server.MapPath("/ThesisReview/admin/upload/review/"&rs("REVIEW_FILE"))
CloseRs rs
reviewtime=Now
If teachtype_id=5 Or teachtype_id=6 Then
	reviewfile_type=2
Else
	reviewfile_type=1
End If
Randomize
filename=toDateTime(reviewtime,1)&Int(Timer)&Int(Rnd()*999)
fullfilename=filename&".pdf"
filepath=Server.MapPath("export")&"\"&fullfilename
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
rg.ExpertMasterLevel=expert_master_level
rg.EvalText=eval_text
rg.CorrelationType=correlation_type
rg.ReviewResult=reviewresult
rg.ReviewLevel=reviewlevel
rg.ThesisType=review_type
If reviewfile_type=2 Then	' ME/MBA评阅书，计算评价指标总分
	rg.Spec=speciality
	rg.Scores=scores
	Dim arrScorePartPower,arrScores,arrScorePower
	Dim scoreParts,partScore,totalScore
	loadReviewScoringInfo review_type,tmp,power1code,power2code
	power1code=Replace(power1code,"[","Array(")
	power1code=Replace(power1code,"]",")")
	power2code=Replace(power2code,"[","Array(")
	power2code=Replace(power2code,"]",")")
	arrScorePartPower=Eval(power1code)
	arrScorePower=Eval(power2code)
	arrScores=Split(scores,",")
	totalScore=0
	k=0
	For i=0 To UBound(arrScorePartPower)
		partScore=0
		For j=0 To UBound(arrScorePower(i))
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
End If
bError=rg.generateReview(filepath,filepath2,filepath3,templatefile,reviewfile_type)=0
Set rg=Nothing

review_result(reviewer_type)=reviewresult
review_level(reviewer_type)=reviewlevel
reviewer_master_level(reviewer_type)=expert_master_level
review_file(reviewer_type)=fullfilename
review_time(reviewer_type)=reviewtime
' 确定处理意见
code=review_result(0)&review_result(1)
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
review_result(2)=finalresult

' 更新记录
sql="SELECT * FROM TEST_THESIS_REVIEW_INFO WHERE ID="&thesisID
GetRecordSet conn,rs,sql,result
rs("REVIEW_RESULT")=join(review_result,",")
rs("REVIEW_LEVEL")=join(review_level,",")
rs("REVIEWER_MASTER_LEVEL")=join(reviewer_master_level,",")
rs("REVIEWER_EVAL"&(reviewer_type+1))=eval_text
rs("REVIEW_FILE")=join(review_file,",")
rs("REVIEWER_EVAL_TIME")=join(review_time,",")
If finalresult<>"6" Then
	rs("REVIEW_STATUS")=rsReviewed
End If
rs.Update()
CloseRs rs
CloseConn conn

updateActiveTime Session("Tid")

logtxt="专家["&expert_name&"]提交评阅意见，论文：《"&subject&"》，作者："&author&"，评阅书："&fullfilename&"。"
writeLog logtxt
%><form id="ret" action="thesisDetail.asp?tid=<%=thesisID%>" method="post">
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
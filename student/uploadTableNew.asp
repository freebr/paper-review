<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<!--#include file="../inc/mail.asp"-->
<!--#include file="common.asp"-->
<!--#include file="tablegen.inc"-->
<%If IsEmpty(Session("Suser")) Then Response.Redirect("../error.asp?timeout")
Dim opr,bOpen,bUpload,bGenerated,filetype
Dim researchway_list
Dim conn,rs,sql,result

bOpen=True
bUpload=True
opr=0
sem_info=getCurrentSemester()
stu_type=Session("StuType")
researchway_list=loadResearchwayList(stu_type)

Connect conn
sql="SELECT *,dbo.getThesisStatusText(1,TASK_PROGRESS,2) AS STAT_TEXT FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE STU_ID="&Session("Stuid")&" ORDER BY PERIOD_ID DESC"
GetRecordSetNoLock conn,rs,sql,result
sql="SELECT * FROM VIEW_STUDENT_INFO WHERE STU_ID="&Session("Stuid")
GetRecordSetNoLock conn,rs2,sql,result
enter_year=rs2("ENTER_YEAR")
tutor_name=rs2("TEACHERNAME")
tutor_duty_name=getProDutyNameOf(rs2("TUTOR_ID"))
If rs.EOF Then
	opr=STUCLI_OPR_TABLE1
	task_progress=tpNone
	str_keywords_ch="''"
	str_keywords_en="''"
Else
	thesisID=rs("ID")
	' 表格审核进度
	task_progress=rs("TASK_PROGRESS")
	Select Case task_progress
	Case tpNone,tpTbl1Uploaded,tpTbl1Unpassed	' 开题报告
		opr=STUCLI_OPR_TABLE1
		bGenerated=task_progress=tpTbl1Uploaded Or task_progress=tpTbl1Unpassed
		filetype=1
	Case tpTbl1Passed,tpTbl2Uploaded,tpTbl2Unpassed	' 中期检查表
		opr=STUCLI_OPR_TABLE2
		bGenerated=task_progress=tpTbl2Uploaded Or task_progress=tpTbl2Unpassed
		filetype=3
	Case tpTbl2Passed,tpTbl3Uploaded,tpTbl3Unpassed	' 预答辩申请表
		opr=STUCLI_OPR_TABLE3
		bGenerated=task_progress=tpTbl3Uploaded Or task_progress=tpTbl3Unpassed
		filetype=5
	Case tpTbl3Passed,tpTbl4Uploaded,tpTbl4Unpassed	' 答辩审批材料
		review_status=rs("REVIEW_STATUS")
		If review_status>=rsReviewEval Then
			opr=STUCLI_OPR_TABLE4
			bGenerated=task_progress=tpTbl4Uploaded Or task_progress=tpTbl4Unpassed
			filetype=7
		Else
			opr=STUCLI_OPR_TABLE3
			bUpload=False
			bRedirectToThesisUpload=True
		End If
	Case tpTbl4Passed ' 答辩审批材料审核通过
		opr=STUCLI_OPR_TABLE4
		bGenerated=True
		filetype=7
		bUpload=False
	Case Else
		bUpload=False
	End Select
	speciality_name=rs("SPECIALITY_NAME")
	subject=rs("THESIS_SUBJECT")
	subject_en=rs("THESIS_SUBJECT_EN")
	research_field=rs("RESEARCHWAY_NAME")
	str_keywords_ch="'"&Join(Split(toPlainString(rs("KEYWORDS")),","),"','")&"'"
	str_keywords_en="'"&Join(Split(toPlainString(rs("KEYWORDS_EN")),","),"','")&"'"
	review_type=rs("REVIEW_TYPE")
	thesis_form=rs("THESIS_FORM")
	tutor_modify_eval=rs("TUTOR_MODIFY_EVAL")
End If
If opr<>0 Then
	bOpen=stuclient.isOpenFor(stu_type,opr)
	startdate=stuclient.getOpentime(opr,STUCLI_OPENTIME_START)
	enddate=stuclient.getOpentime(opr,STUCLI_OPENTIME_END)
	If Not bOpen Then bUpload=False
	If opr<=STUCLI_OPR_TABLE3 Then
		bTblThesisUploaded=Not IsNull(rs("TBL_THESIS_FILE"&opr))
	End If
End If
' 确定表格模板文件名
Select Case opr
Case STUCLI_OPR_TABLE1:template_name="ktbg"
Case STUCLI_OPR_TABLE2:template_name="zqjcb"
Case STUCLI_OPR_TABLE3:template_name="ydbyjs"
Case STUCLI_OPR_TABLE4:template_name="spcl"
End Select
If opr=STUCLI_OPR_TABLE4 Then
	Select Case stu_type
	Case 5
		template_name=template_name&"_me"
	Case 6
		template_name=template_name&"_mba"
	Case 7
		template_name=template_name&"_emba"
	Case 9
		template_name=template_name&"_mpacc"
	End Select
End If
template_file="template/doc/"&template_name&".doc"

curStep=Request.QueryString("step")
Select Case curStep
Case vbNullstring ' 填写信息页面
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/student.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
<script src="../scripts/utils.js" type="text/javascript"></script>
<script src="../scripts/upload.js" type="text/javascript"></script>
<script src="../scripts/uploadTable.js" type="text/javascript"></script>
<script src="../scripts/keywordList.js" type="text/javascript"></script>
<meta name="theme-color" content="#2D79B2" />
<title>在线填写表格</title>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>在线填写表格</b></font
<table class="tblform" width="1000"><tr><td class="summary"><p><%
	If Not bOpen Then
%><span class="tip">提交<%=arrStuOprName(opr)%>的时间为<%=toDateTime(startdate,1)%>至<%=toDateTime(enddate,1)%>，本专业提交通道已关闭或当前不在开放时间内，不能提交表格！</span><%
	ElseIf Not bUpload Then
%><span class="tip">当前状态为【<%=rs("STAT_TEXT")%>】，不能提交表格！</span><%
	Else
%>当前填写的是：<span style="color:#ff0000;font-weight:bold"><%=arrStuOprName(opr)%></span><br/>
请填写以下信息，然后点击&quot;提交&quot;按钮：<%
	End If %></p></td></tr>
<tr><td align="center"><form id="fmThesis" action="?step=1" method="post">
<input type="hidden" name="stuid" value="<%=Session("Stuid")%>" />
<table class="tblform">
<!--<tr><td><span class="tip">以下信息均为必填项</span></td></tr>-->
<tr><td align="center"><%
	If bUpload Then
		Select Case opr
		Case STUCLI_OPR_TABLE1
%><!--#include file="template/form_ktbg.html"--><%
		Case STUCLI_OPR_TABLE2
%><!--#include file="template/form_zqjcb.html"--><%
		Case STUCLI_OPR_TABLE3
%><!--#include file="template/form_ydbyjs.html"--><%
		Case STUCLI_OPR_TABLE4
%><!--#include file="template/form_spcl.html"--><%
		End Select
	End If %>
</td></tr><%
	If opr>0 And opr<=STUCLI_OPR_TABLE3 And Not bTblThesisUploaded Then %>
<tr><td align="center"><span class="tip">提示：您目前尚未上传<%=arrTblThesis(opr)%>，<a href="uploadTableThesis.asp">点击这里上传。</a></span></td></tr><%
	End If %>
<tr><td align="center"><p><%
	If bUpload Then
%><input type="submit" name="btnsubmit" value="提 交" />&nbsp;<%
	End If
	If bGenerated Then
%><input type="button" id="btndownload" value="下载打印" />&nbsp;<%
	End If
	If opr<>STUCLI_OPR_TABLE4 Then
%><input type="button" id="btnuploadtblthesis" value="上传<%=arrTblThesis(opr)%>" />&nbsp;<%
	End If
%><input type="button" id="btnreturn" value="返回首页" onclick="location.href='default.asp'" /></p></td></tr>
<tr><td><%
	If opr<>0 Then %>
<div style="text-align:right"><hr />
<a href="<%=template_file%>" target="_blank"><img src="../images/down.png" />下载<%=arrStuOprName(opr)%>模板...</a></div><%
	End If %></td></tr></table></form></td></tr></table></center>
<script type="text/javascript">
<%
	If opr=STUCLI_OPR_TABLE1 And (stu_type=5 Or stu_type=6) Then %>
		initResearchFieldSelectBox($('#research_field_select'),<%=stu_type%>);
		$('#school_tutor_research_field_select').change(function(){
			$('input[name="school_tutor_research_field"]').val(this.options[this.selectedIndex].innerText);
		});
		$('#research_field_select').change(function(){
			initSubResearchFieldSelectBox($('#school_tutor_research_field_select'),$(this),this.selectedIndex);
			$('#school_tutor_research_field_select').change();
			$('input[name="research_field"]').val(this.options[this.selectedIndex].innerText);
		});<%
	End If %>
	$('form').submit(function() {
			return submitUploadForm(this);
		}).find(':submit').attr('disabled',<%=LCase(Not bUpload)%>);
	$(':button#btnuploadtblthesis').click(
		function() {
			window.location.href='uploadTableThesis.asp';
		});
	$(':button#btndownload').click(
		function() {
			window.location.href='fetchfile.asp?tid=<%=thesisID%>&type=<%=filetype%>';
		});
</script></body></html><%
Case 1	' 上传进程

	If Not bOpen Then
		bError=True
		errdesc="提交"&arrStuOprName(opr)&"的时间为"&FormatDateTime(startdate,1)&"至"&FormatDateTime(enddate,1)&"，本专业提交通道已关闭或当前不在开放时间内，不能提交表格！"
	ElseIf Not bUpload Then
		bError=True
		errdesc="当前状态为【"&rs("STAT_TEXT")&"】，不能提交表格！"
	End If
	If bError Then
		%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
  	CloseRs rs
  	CloseRs rs2
  	CloseConn conn
		Response.End
	End If
	
	Dim fso
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	
	' 检查上传目录是否存在
	strUploadPath=Server.MapPath("upload")
	If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)
	Set fso=Nothing
	' 生成表格文件名
	fileid=FormatDateTime(Now(),1)&Int(Timer)
	strDestTableFile=fileid&".doc"
	strDestTablePath=strUploadPath&"\"&strDestTableFile
	' 生成表格
	Dim tg:Set tg=New TableGen
	
	Select Case opr
	Case STUCLI_OPR_TABLE1
	
		subject=Request.Form("subject_ch")
		subject_en=Request.Form("subject_en")
		research_field=Request.Form("research_field")
		school_tutor_name=Request.Form("school_tutor_name")
		school_tutor_duty=Request.Form("school_tutor_duty")
		school_tutor_research_field=Request.Form("school_tutor_research_field")
		afterschool_tutor_name=Request.Form("afterschool_tutor_name")
		afterschool_tutor_duty=Request.Form("afterschool_tutor_duty")
		afterschool_tutor_expertise=Request.Form("afterschool_tutor_expertise")
		issue_source=Request.Form("issue_source")
		abstract=Request.Form("abstract")
		research_background=Request.Form("research_background")
		research_solution=Request.Form("research_solution")
		anticipated_result=Request.Form("anticipated_result")
		
		tg.addInfo "StuName",Session("StuName")
		tg.addInfo "StuNo",Session("StuNo")
		tg.addInfo "StuType",stu_type
		tg.addInfo "ResearchField",research_field
		
		tg.addInfo "ThesisSubjectCh",subject
		tg.addInfo "ThesisSubjectEn",subject_en
		tg.addInfo "SchoolTutorName",school_tutor_name
		tg.addInfo "SchoolTutorDuty",school_tutor_duty
		tg.addInfo "SchoolTutorResearchField",school_tutor_research_field
		tg.addInfo "AfterSchoolTutorName",afterschool_tutor_name
		tg.addInfo "AfterSchoolTutorDuty",afterschool_tutor_duty
		tg.addInfo "AfterSchoolTutorExpertise",afterschool_tutor_expertise
		tg.addInfo "IssueSource",issue_source
		tg.addInfo "Abstract",abstract
		For i=1 To Request.Form("keyword_ch").Count
			If Len(Trim(Request.Form("keyword_ch")(i))) Then
				If i>1 Then keywords_ch=keywords_ch&"　"
				keywords_ch=keywords_ch&Trim(Request.Form("keyword_ch")(i))
			End If
		Next
		tg.addInfo "KeywordsCh",keywords_ch
		For i=1 To Request.Form("keyword_en").Count
			If Len(Trim(Request.Form("keyword_en")(i))) Then
				If i>1 Then keywords_en=keywords_en&"/"
				keywords_en=keywords_en&Trim(Request.Form("keyword_en")(i))
			End If
		Next
		tg.addInfo "KeywordsEn",keywords_en
		
		tg.addInfo "ResearchBackground",research_background
		tg.addInfo "ResearchSolution",research_solution
		
		addFormInfoToArray tg,work_schedule_duration,"work_schedule_duration","WorkScheduleDuration"
		addFormInfoToArray tg,work_schedule_content,"work_schedule_content","WorkScheduleContent"
		addFormInfoToArray tg,work_schedule_memo,"work_schedule_memo","WorkScheduleMemo"
		
		tg.addInfo "AnticipatedResult",anticipated_result
		
	Case STUCLI_OPR_TABLE2
	
		subject=Request.Form("subject")
		research_field=Request.Form("research_field")
		thesis_progress=Request.Form("thesis_progress")
		work_schedule=Request.Form("work_schedule")
		
		tg.addInfo "StuName",Session("StuName")
		tg.addInfo "StuNo",Session("StuNo")
		tg.addInfo "StuType",stu_type
		tg.addInfo "ResearchField",research_field
		
		tg.addInfo "ThesisSubject",subject
		tg.addInfo "ThesisProgress",thesis_progress
		tg.addInfo "WorkSchedule",work_schedule
		
	Case STUCLI_OPR_TABLE3
	
		grade=Request.Form("grade")
		speciality_name=Request.Form("speciality_name")
		subject=Request.Form("subject")
		predefence_date=Request.Form("predefence_date")
		
		If Len(grade)<>4 Or Not IsNumeric(grade) Then
			bError=True
			errdesc="年级填写无效，请重新输入（格式为四位数字）！"
		ElseIf Not IsDate(predefence_date) Then
			bError=True
			errdesc="预答辩日期填写无效，请重新输入！"
		End If
		If bError Then
	%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
			CloseRs rs
	  	CloseRs rs2
	  	CloseConn conn
			Response.End
		End If
		
		tg.addInfo "StuName",Session("StuName")
		tg.addInfo "Grade",grade
		tg.addInfo "SpecialityName",speciality_name
		tg.addInfo "ThesisSubject",subject
		tg.addInfo "PredefenceYear",Year(predefence_date)
		tg.addInfo "PredefenceMonth",Month(predefence_date)
		tg.addInfo "PredefenceDay",Day(predefence_date)
		
	Case STUCLI_OPR_TABLE4
	
		research_field=Request.Form("research_field")
		school_tutor_name_duty=Request.Form("school_tutor_name_duty")
		after_school_tutor_name_duty=Request.Form("after_school_tutor_name_duty")
		degree_application=Request.Form("degree_application")
		sex=Request.Form("sex")
		birth_ym=toYearMonth(Request.Form("birth_ym")(1),Request.Form("birth_ym")(2))
		idcard_no=Request.Form("idcard_no")
		native_place=Request.Form("native_place")
		nation=Request.Form("nation")
		entrance_ym=toYearMonth(Request.Form("entrance_ym")(1),Request.Form("entrance_ym")(2))
		political_status=Request.Form("political_status")
		study_type=Request.Form("study_type")
		workplace_job=Request.Form("workplace_job")
		graduated_at=Request.Form("graduated_at")
		speciality_name=Request.Form("speciality_name")
		graduation_ym=toYearMonth(Request.Form("graduation_ym")(1),Request.Form("graduation_ym")(2))
		last_degree=Request.Form("last_degree")
		honor_penalty=Request.Form("honor_penalty")
		subject=Request.Form("subject")
		thesis_word_count=Request.Form("thesis_word_count")
		issue_source=Request.Form("issue_source")
		speciality_name_code=Request.Form("speciality_name_code")
		thesis_type=Request.Form("thesis_type")
		thesis_duration=Request.Form("thesis_duration")
		thesis_form=Request.Form("thesis_form")
		thesis_introduction=Request.Form("thesis_introduction")
		thesis_count=Request.Form("thesis_count")
		thesis_count_domestic_journal=Request.Form("thesis_count_domestic_journal")
		thesis_count_domestic_congress=Request.Form("thesis_count_domestic_congress")
		thesis_count_foreign_journal=Request.Form("thesis_count_foreign_journal")
		thesis_count_overseas_congress=Request.Form("thesis_count_overseas_congress")
		thesis_count_patent=Request.Form("thesis_count_patent")
		embodied_count=Request.Form("embodied_count")
		school_tutor_eval=Request.Form("school_tutor_eval")
		
		tg.addInfo "StuNo",Session("StuNo")
		tg.addInfo "StuName",Session("StuName")
		tg.addInfo "ResearchField",research_field
		tg.addInfo "SchoolTutorNameDuty",school_tutor_name_duty
		tg.addInfo "AfterSchoolTutorNameDuty",after_school_tutor_name_duty
		tg.addInfo "FillInDate",Year(Now)&"年"&Month(Now)&"月"&Day(Now)&"日"
		tg.addInfo "DegreeApplication",degree_application
		tg.addInfo "Sex",sex
		tg.addInfo "BirthYearMonth",birth_ym
		tg.addInfo "IDCardNo",idcard_no
		tg.addInfo "NativePlace",native_place
		tg.addInfo "Nation",nation
		tg.addInfo "EntranceYearMonth",entrance_ym
		tg.addInfo "PoliticalStatus",political_status
		tg.addInfo "StudyType",study_type
		tg.addInfo "WorkplaceJob",workplace_job
		tg.addInfo "GraduatedAt",graduated_at
		tg.addInfo "SpecialityName",speciality_name
		tg.addInfo "GraduationYearMonth",graduation_ym
		tg.addInfo "LastDegree",last_degree
		
		addFormInfoToArray tg,resume_duration,"resume_duration","ResumeDuration"
		addFormInfoToArray tg,resume_place,"resume_place","ResumePlace"
		addFormInfoToArray tg,resume_job,"resume_job","ResumeJob"
		tg.addInfo "HonorPenalty",honor_penalty
		
		tg.addInfo "ThesisSubject",subject
		tg.addInfo "ThesisWordCount",thesis_word_count
		tg.addInfo "IssueSource",issue_source
		tg.addInfo "SpecialityNameCode",speciality_name_code
		tg.addInfo "ThesisType",thesis_type
		tg.addInfo "ThesisDuration",thesis_duration
		tg.addInfo "ThesisForm",thesis_form
		tg.addInfo "ThesisIntroduction",thesis_introduction
		
		addFormInfoToArray tg,achievement_name,"achievement_name","AchievementName"
		addFormInfoToArray tg,achievement_ym,"achievement_ym","AchievementYearMonth"
		addFormInfoToArray tg,achievement_department,"achievement_department","AchievementDepartment"
		addFormInfoToArray tg,achievement_authornum,"achievement_authornum","AchievementAuthorNum"
		addFormInfoToArray tg,achievement_embody,"achievement_embody","AchievementEmbody"
		ReDim achievement_id(UBound(achievement_name))
		For i=0 To UBound(achievement_id)
			If Len(Trim(achievement_name(i)))=0 Then Exit For
			achievement_id(i)=i+1
		Next
		tg.addInfo "AchievementID",achievement_id
		
		tg.addInfo "ThesisCount",thesis_count
		tg.addInfo "ThesisCountDomesticJournal",thesis_count_domestic_journal
		tg.addInfo "ThesisCountDomesticCongress",thesis_count_domestic_congress
		tg.addInfo "ThesisCountForeignJournal",thesis_count_foreign_journal
		tg.addInfo "ThesisCountOverseasCongress",thesis_count_overseas_congress
		tg.addInfo "ThesisCountPatent",thesis_count_patent
		tg.addInfo "EmbodiedCount",embodied_count
		tg.addInfo "SchoolTutorEval",school_tutor_eval
		
	End Select
	
	tg.generateTable strDestTablePath,template_name
	Set tg=Nothing
	
	Dim arrTableFieldName,arrNewTaskProgress
	arrTableFieldName=Array("","TABLE_FILE1","TABLE_FILE2","TABLE_FILE3","TABLE_FILE4")
	arrNewTaskProgress=Array(0,tpTbl1Uploaded,tpTbl2Uploaded,tpTbl3Uploaded,tpTbl4Uploaded)
	' 关联到数据库
	sql="SELECT * FROM TEST_THESIS_REVIEW_INFO WHERE STU_ID="&Session("StuId")&" ORDER BY PERIOD_ID DESC"
	GetRecordSet conn,rs3,sql,result
	If rs3.EOF Then
		' 添加记录
		rs3.AddNew()
	End If
	If opr=STUCLI_OPR_TABLE1 Then	' 开题报告，录入论文基本信息
		rs3("STU_ID")=Session("StuId")
		rs3("REVIEW_TYPE")=new_review_type
		rs3("REVIEW_STATUS")=rsNone
		rs3("REVIEW_FILE_STATUS")=0
		rs3("REVIEW_RESULT")="5,5,6"
		rs3("REVIEW_LEVEL")="0,0"
	End If
	rs3("PERIOD_ID")=sem_info(3)
	rs3("THESIS_SUBJECT")=subject
	rs3("THESIS_SUBJECT_EN")=subject_en
	rs3("KEYWORDS")=Replace(Request.Form("keyword_ch"),", ",",")
	rs3("KEYWORDS_EN")=Replace(Request.Form("keyword_en"),", ",",")
	rs3("RESEARCHWAY_NAME")=research_field
	rs3(arrTableFieldName(opr))=strDestTableFile
	rs3("TASK_PROGRESS")=arrNewTaskProgress(opr)
	rs3.Update()
	CloseRs rs3
	
	Dim logtxt
	logtxt="学生["&Session("StuName")&"]提交["&arrStuOprName(opr)&"]。"
	WriteLogForReviewSystem logtxt
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>提交论文表格</title>
<link href="../css/student.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite"><%
	If Not bError Then %>
<form id="fmFinish" action="default.asp" method="post">
<input type="hidden" name="filename" value="<%=strDestTableFile%>" />
</form>
<script type="text/javascript">alert("提交成功！");$('#fmFinish').submit();</script><%
	Else
%><script type="text/javascript">alert("<%=errdesc%>");history.go(-1);</script><%
	End If
%></body></html><%
End Select
CloseRs rs2
CloseRs rs
CloseConn conn
%>
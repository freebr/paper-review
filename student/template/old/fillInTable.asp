<!--#include file="../inc/automation/PaperFormWriter.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("StuId")) Then Response.Redirect("../error.asp?timeout")
Dim is_new_dissertation:is_new_dissertation=False
Dim activity_id,section_id,time_flag,uploadable,is_generated,filetype
Dim researchway_list
Dim conn,rs,sql,count

activity_id=0
section_id=0
uploadable=False
stu_type=Session("StuType")
researchway_list=loadResearchwayList(stu_type)

ConnectDb conn
sql="SELECT *,dbo.getThesisStatusText(1,TASK_PROGRESS,2) AS STAT_TEXT FROM ViewDissertations WHERE STU_ID="&Session("Stuid")&" ORDER BY ActivityId DESC"
GetRecordSetNoLock conn,rs,sql,count
sql="SELECT STU_NO,SEX,CLASS_NAME,ENTER_YEAR,TEACHERNAME,TUTOR_ID FROM ViewStudentInfo WHERE STU_ID="&Session("Stuid")
GetRecordSetNoLock conn,rsStu,sql,count
stu_no=rsStu("STU_NO")
stu_sex=Abs(Int(rsStu("SEX")="女"))
foreign_student=InStr(rsStu("CLASS_NAME"),"Class") > 0
If foreign_student Then
	stu_type_name=dictStuTypes(stu_type)(1)
Else
	stu_type_name=dictStuTypes(stu_type)(0)
End If
enter_year=rsStu("ENTER_YEAR")
tutor_name=rsStu("TEACHERNAME")
tutor_name_en=getPinyinOfName(rsStu("TEACHERNAME"))
tutor_duty_name=getProDutyNameOf(rsStu("TUTOR_ID"))
If rs.EOF Then
	section_id=sectionUploadKtbgb
	task_progress=tpNone
	str_keywords_ch="''"
	str_keywords_en="''"
Else
	paper_id=rs("ID")
	activity_id=rs("ActivityId")
	' 表格审核进度
	task_progress=rs("TASK_PROGRESS")
	Select Case task_progress
	Case tpNone,tpTbl1Uploaded,tpTbl1Unpassed	' 开题报告
		section_id=sectionUploadKtbgb
		is_generated=Not IsNull(rs("TABLE_FILE1"))
		filetype=1
	Case tpTbl1Passed,tpTbl2Uploaded,tpTbl2Unpassed	' 中期考核表
		section_id=sectionUploadZqkhb
		is_generated=Not IsNull(rs("TABLE_FILE2"))
		filetype=3
	Case tpTbl2Passed,tpTbl3Uploaded,tpTbl3Unpassed	' 预答辩意见书
		section_id=sectionUploadYdbyjs
		is_generated=Not IsNull(rs("TABLE_FILE3"))
		filetype=5
	Case tpTbl3Passed,tpTbl4Uploaded,tpTbl4Unpassed	' 答辩审批材料
		review_status=rs("REVIEW_STATUS")
		If review_status>=rsReviewEval Then
			section_id=sectionUploadSpclb
			is_generated=Not IsNull(rs("TABLE_FILE4"))
			' 获取导师对答辩论文的审核意见，自动填充到指导教师评语
			Dim audit_info:audit_info = getAuditInfo(paper_id, Null, auditTypeDefence)
			If audit_info(0)("Id") <> "0" Then
				tutor_eval = audit_info(0)("Comment")
			End If
			filetype=7
		End If
	Case tpTbl4Passed ' 答辩审批材料审核通过
		is_generated=True
		filetype=7
	End Select
	speciality_name=rs("SPECIALITY_NAME")
	If foreign_student Then
		If speciality_name="工商管理硕士" Then
			speciality_name="Master of Business Administration"
		End If
	End If
	subject=rs("THESIS_SUBJECT")
	subject_en=rs("THESIS_SUBJECT_EN")
	sub_research_field=rs("RESEARCHWAY_NAME")
	str_keywords_ch="'"&Join(Split(toPlainString(rs("KEYWORDS")),"；"),"','")&"'"
	str_keywords_en="'"&Join(Split(toPlainString(rs("KEYWORDS_EN")),"；"),"','")&"'"
	review_type=rs("REVIEW_TYPE")
	thesis_form=rs("THESIS_FORM")
	tutor_modify_eval=rs("TUTOR_MODIFY_EVAL")
End If
If section_id<>0 Then
	' 开题报告（EMBA为预答辩意见书），录入论文基本信息
	is_new_dissertation=section_id=sectionUploadKtbgb Or stu_type=7 And section_id=sectionUploadYdbyjs
	If rs.EOF Then
		uploadable=True
	ElseIf Not isActivityOpen(rs("ActivityId")) Then
		time_flag=-3
	Else
		Set current_section=getSectionInfo(rs("ActivityId"), stu_type, section_id)
		time_flag=compareNowWithSectionTime(current_section)
		uploadable=time_flag=0
		If section_id<=sectionUploadYdbyjs Then
			is_tbl_thesis_uploaded=Not IsNull(rs("TBL_THESIS_FILE"&section_id))
		End If
	End If
End If
' 确定表格模板文件名
Select Case section_id
Case sectionUploadKtbgb:template_name="ktbgb"
Case sectionUploadZqkhb:template_name="zqkhb"
Case sectionUploadYdbyjs:template_name="ydbyjs"
Case sectionUploadSpclb
	If foreign_student Then
		template_name="spclb_en"
	Else
		template_name="spclb"
	End If
End Select
Select Case stu_type
Case 5
	prefix="me_"
Case 6
	prefix="mba_"
Case 7
	prefix="emba_"
Case 9
	prefix="mpacc_"
End Select
version="20200814"
template_file=resolvePath("template/doc",version,prefix&template_name&".doc")

view_name = "fillInTable_"&arrStuOprName(section_id)
' 获取视图状态
view_state = getViewState(Session("StuId"),usertypeStudent,view_name)

step=Request.QueryString("step")
Select Case step
Case vbNullstring ' 填写信息页面
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>在线填写表格</title>
<% useStylesheet "student", "jeasyui" %>
<% useScript "jquery", "jeasyui", "common", "upload", "fillInTable", "keywordList", "*viewState" %>
</head>
<body>
<center><font size=4><b>在线填写表格</b></font>
<form id="fmTable" action="?step=1" method="post">
<table class="form" width="1000"><tr><td class="summary"><%
	If Not uploadable Then
		If time_flag=-3 Then
%><p><span class="tip">当前评阅活动【<%=rs("ActivityName")%>】已关闭，不能提交表格！</span></p><%
		ElseIf time_flag=-2 Then
%><p><span class="tip">【<%=current_section("Name")%>】环节已关闭，不能提交表格！</span></p><%
		ElseIf time_flag<>0 Then
%><p><span class="tip">【<%=current_section("Name")%>】环节开放时间为<%=toDateTime(current_section("StartTime"),1)%>至<%=toDateTime(current_section("EndTime"),1)%>，当前不在开放时间内，不能提交表格！</span></p><%
		Else
%><p><span class="tip">当前状态为【<%=rs("STAT_TEXT")%>】，没有需提交的表格！</span></p><%
		End If
	Else
%><p>当前填写的是：<span style="color:#ff0000;font-weight:bold"><%=arrStuOprName(section_id)%></span></p><%
		If is_new_dissertation Then %>
<p>请选择您要参加的评阅活动：
<input id="activity_id" class="easyui-combobox" name="activity_id"
    data-options="valueField: 'id',
	textField: 'name',
	editable: false,
	prompt: '【请选择】',
	width: 300,
	panelHeight: 100,<%
	If activity_id<>0 Then %>
	value: <%=activity_id%>,<%
	End If %>
	url: '../api/get-attendable-activities',
	loadFilter: Common.curryLoadFilter(Array.prototype.reverse)"></p><%
		End If %>
<p>请填写以下信息，确认无误后点击&quot;提交&quot;按钮生成表格：</p><%
	End If %></td></tr>
<tr><td align="center">
<table class="form">
<!--<tr><td><span class="tip">以下信息均为必填项</span></td></tr>-->
<tr><td align="center"><%
	If uploadable Then
		Select Case section_id
		Case sectionUploadKtbgb
%><!--#include file="template/form_ktbg.html"--><%
		Case sectionUploadZqkhb
%><!--#include file="template/form_zqkhb.html"--><%
		Case sectionUploadYdbyjs
%><!--#include file="template/form_ydbyjs.html"--><%
		Case sectionUploadSpclb
			If foreign_student Then
%><!--#include file="template/form_spcl_en.html"--><%
			Else
%><!--#include file="template/form_spcl.html"--><%
			End If
		End Select
	End If %>
</td></tr><%
	If section_id>0 And section_id<=sectionUploadYdbyjs And uploadable And Not is_tbl_thesis_uploaded Then %>
<tr><td align="center"><span class="tip">提示：您目前尚未上传<%=arrTblThesis(section_id)%>，<a href="uploadTablePaper.asp">点击这里上传。</a></span></td></tr><%
	End If %>
<tr><td align="center"><p><%
	If uploadable Then
%><input type="button" id="btnsavedraft" value="保存草稿" />&nbsp;
<input type="button" id="btnloaddraft" value="读取草稿" />&nbsp;
<input type="submit" id="btnsubmit" value="提 交" />&nbsp;<%
	End If
	If is_generated Then
%><input type="button" id="btndownload" value="下载打印已提交表格" />&nbsp;<%
	End If
	If uploadable And section_id<>sectionUploadSpclb Then
%><input type="button" id="btnuploadtblthesis" value="上传<%=arrTblThesis(section_id)%>" />&nbsp;<%
	End If
%><input type="button" id="btnreturn" value="返回首页" onclick="location.href='home.asp'" /></p></td></tr>
<tr><td><%
	If section_id<>0 Then %>
<div style="text-align:right"><hr />
<a href="<%=template_file%>" target="_blank"><img src="../images/down.png" />下载<%=arrStuOprName(section_id)%>模板...</a></div><%
	End If %></td></tr></table></td></tr></table></form></center>
<script type="text/javascript">
	function initCurrentViewState() {
		initViewState($("form"), {
			user_id: <%=Session("StuId")%>,
			user_type: <%=usertypeStudent%>,
			view_name: "<%=view_name%>",
			view_state: <%=isNullString(view_state, "null")%>
		}, function(data) {
			if (!data.keyword_ch) return;
			var keywords_ch = [], keywords_en = [];
			[].concat(data.keyword_ch).forEach(function(item) {
				keywords_ch.push(item.value);
			});
			[].concat(data.keyword_en).forEach(function(item) {
				keywords_en.push(item.value);
			});
			setKeywords(keywords_ch, keywords_en);
		});
	}<%
	If section_id=sectionUploadKtbgb And uploadable Then %>
	$('select[name="sub_research_field_select"]').change(function(){
		$('input[name="sub_research_field"]').val(!this.value.length?'':$(this).find('option:selected').text());
		var $custom_field=$('input[name="custom_sub_research_field"]');
		if(this.value=='-1')
			$custom_field.show().focus();
		else
			$custom_field.hide();
	});
	$('select[name="school_tutor_research_field_select"]').change(function(){
		$('input[name="school_tutor_research_field"]').val(!this.value.length?'':$(this).find('option:selected').text());
	});
	$('select[name="research_field_select"]').change(function(){
		initSubResearchFieldSelectBox($('select[name="sub_research_field_select"]'),$(this),this.value, '<%=sub_research_field%>');
		initSubResearchFieldSelectBox($('select[name="school_tutor_research_field_select"]'),$(this),this.value, '<%=sub_research_field%>');
		$('select[name="sub_research_field_select"]').change();
		$('select[name="school_tutor_research_field_select"]').change();
		$('input[name="research_field"]').val(!this.value.length?'':$(this).find('option:selected').text());
	});
	initResearchFieldSelectBox($('select[name="research_field_select"]'),<%=stu_type%>)
		.then(initCurrentViewState);<%
	Else
	%>
		initCurrentViewState();<%
	End If %>
	$('form').submit(function(event) {<%
	If section_id=sectionUploadKtbgb And uploadable Then %>
		if(!checkKeywords()) {
			event.preventDefault();
			return false;
		}<%
	End If %>
		return submitUploadForm(this);
	});
	$(':button#btnuploadtblthesis').click(
		function() {
			window.location.href='uploadTablePaper.asp';
		}
	);
	$(':button#btndownload').click(
		function() {
			window.location.href='fetchDocument.asp?tid=<%=paper_id%>&type=<%=filetype%>';
		}
	);
</script></body></html><%
Case 1	' 上传进程

	If time_flag=-3 Then
		bError=True
		errMsg=Format("当前评阅活动【{0}】已关闭，不能提交表格！", rs("ActivityName"))
	ElseIf time_flag=-2 Then
		bError=True
		errMsg=Format("【{0}】环节已关闭，不能提交表格！", current_section("Name"))
	ElseIf time_flag<>0 Then
		bError=True
		errMsg=Format("【{0}】环节开放时间为{1}至{2}，当前不在开放时间内，不能提交表格！",_
			current_section("Name"),_
			toDateTime(current_section("StartTime"),1),_
			toDateTime(current_section("EndTime"),1))
	ElseIf Not uploadable Then
		bError=True
		errMsg="当前状态为【"&rs("STAT_TEXT")&"】，没有需提交的表格！"
	End If
	If bError Then
		CloseRs rs
		CloseRs rsStu
		CloseConn conn
		showErrorPage errMsg, "提示"
	End If

	Dim fso:Set fso=CreateFSO()

	' 检查上传目录是否存在
	upload_path=Server.MapPath("upload")
	If Not fso.FolderExists(upload_path) Then fso.CreateFolder(upload_path)
	Set fso=Nothing
	' 生成表格文件名
	export_file_name=timestamp()&".doc"
	export_file_path=upload_path&"\"&export_file_name
	
	Dim params, arr, i
	' 生成表格
	Dim tg:Set tg=New PaperFormWriter

	Select Case section_id
	Case sectionUploadKtbgb

		Set params=getFormParams("activity_id","subject","subject_en","research_field_select","research_field",_
			"sub_research_field_select","sub_research_field",_
			"custom_sub_research_field","school_tutor_name","school_tutor_duty",_
			"school_tutor_research_field_select","school_tutor_research_field",_
			"afterschool_tutor_name","afterschool_tutor_duty","afterschool_tutor_expertise",_
			"issue_source","abstract","keyword_ch_all","keyword_en_all","research_background",_
			"research_solution","work_schedule_duration","work_schedule_content",_
			"work_schedule_memo","anticipated_result","view_state")

		If Len(params("activity_id"))=0 Then
			bError=True
			errMsg="请选择要参加的评阅活动！"
		ElseIf stu_type = 5 And Len(params("research_field_select"))=0 Then
			bError=True
			errMsg="请选择工程领域！"
		ElseIf Len(params("sub_research_field_select"))=0 Then
		 	bError=True
		 	errMsg="请选择研究方向！"
		ElseIf params("sub_research_field_select")="-1" And Len(params("custom_sub_research_field"))=0 Then
		 	params("custom_sub_research_field")="其他"
		End If
		If bError Then
			CloseRs rs
			CloseRs rsStu
			CloseConn conn
			showErrorPage errMsg, "提示"
		End If

		tg.setField "StuName",Session("StuName")
		tg.setField "StuNo",Session("StuNo")
		tg.setField "StuType",stu_type
		tg.setField "ResearchField",params("research_field")

		tg.setField "ThesisSubjectCh",params("subject")
		tg.setField "ThesisSubjectEn",params("subject_en")
		tg.setField "SchoolTutorName",params("school_tutor_name")
		tg.setField "SchoolTutorDuty",params("school_tutor_duty")
		tg.setField "SchoolTutorResearchField",params("school_tutor_research_field")
		tg.setField "AfterSchoolTutorName",params("afterschool_tutor_name")
		tg.setField "AfterSchoolTutorDuty",params("afterschool_tutor_duty")
		tg.setField "AfterSchoolTutorExpertise",params("afterschool_tutor_expertise")
		tg.setField "IssueSource",params("issue_source")
		tg.setField "Abstract",params("abstract")
		tg.setField "KeywordsCh",Replace(params("keyword_ch_all"),", ","　")
		tg.setField "KeywordsEn",Replace(params("keyword_en_all"),", ","/")

		tg.setField "ResearchBackground",params("research_background")
		tg.setField "ResearchSolution",params("research_solution")

		tg.setArray "WorkScheduleDuration",params("work_schedule_duration")
		tg.setArray "WorkScheduleContent",params("work_schedule_content")
		tg.setArray "WorkScheduleMemo",params("work_schedule_memo")

		tg.setField "AnticipatedResult",params("anticipated_result")

	Case sectionUploadZqkhb

		Set params=getFormParams("subject","research_field","thesis_progress","work_schedule","view_state")

		tg.setField "StuName",Session("StuName")
		tg.setField "StuNo",Session("StuNo")
		tg.setField "StuType",stu_type
		tg.setField "ResearchField",params("research_field")

		tg.setField "ThesisSubject",params("subject")
		tg.setField "ThesisProgress",params("thesis_progress")
		tg.setField "WorkSchedule",params("work_schedule")

	Case sectionUploadYdbyjs

		Set params=getFormParams("activity_id","grade","speciality_name","subject","predefence_date","view_state")

		If is_new_dissertation And Len(params("activity_id"))=0 Then
			bError=True
			errMsg="请选择要参加的评阅活动！"
		ElseIf Not IsNumeric(params("grade")) Or Len(params("grade"))<>4 Then
			bError=True
			errMsg="年级填写无效，请重新输入（格式为四位数字）！"
		ElseIf Not IsDate(params("predefence_date")) Then
			bError=True
			errMsg="预答辩日期填写无效，请重新输入！"
		End If
		If bError Then
			CloseRs rs
			CloseRs rsStu
			CloseConn conn
			showErrorPage errMsg, "提示"
		End If

		tg.setField "StuName",Session("StuName")
		tg.setField "Grade",params("grade")
		tg.setField "SpecialityName",params("speciality_name")
		tg.setField "ThesisSubject",params("subject")
		tg.setField "PredefenceYear",Year(params("predefence_date"))
		tg.setField "PredefenceMonth",Month(params("predefence_date"))
		tg.setField "PredefenceDay",Day(params("predefence_date"))

	Case sectionUploadSpclb

		Set params=getFormParams("degree_application","sex","ethnic","nation","political_status","idcard_no","speciality_name",_
			"tutor_name","research_field","study_type","workplace_job","graduated_at","before_speciality_name","last_degree",_
			"resume_duration","resume_place","resume_job",_
			"honor_penalty","achievement_name","achievement_ym","achievement_source","achievement_authornum","achievement_status",_
			"dissertation_subject","dissertation_keywords","dissertation_word_count","issue_source","project_name_code",_
			"dissertation_type","dissertation_duration","tutor_eval","view_state")
		debug("count="&Request.Form("birthday").Count)
		debug(",value="&Request.Form("birthday"))
		params("birthday")=toYearMonthDate(Request.Form("birthday")(1),Request.Form("birthday")(2),Request.Form("birthday")(3))
		params("entrance_ym")=toYearMonth(Request.Form("entrance_ym")(1),Request.Form("entrance_ym")(2))
		params("graduation_ym")=toYearMonth(Request.Form("graduation_ym")(1),Request.Form("graduation_ym")(2))
		ReDim arr(params("resume_place").Count-1)
		For i=0 To UBound(arr)
				If Len(params("resume_place")(i+1))=0 Then
				arr(i)=""
			Else
				arr(i)=Replace(params("resume_duration")(i+1), "_", "")
			End If
		Next
		params("resume_duration")=arr

		If Len(params("sex"))=0 Then
			showErrorPage "请完善【性别】信息！", "提示"
		ElseIf stu_type=5 And Len(params("research_field"))=0 Then
			showErrorPage "请完善【工程领域】信息！", "提示"
		ElseIf Len(params("issue_source"))=0 Then
			showErrorPage "请完善【论文选题来源】信息！", "提示"
		ElseIf Len(params("dissertation_type"))=0 Then
			showErrorPage "请完善【论文类型】信息！", "提示"
		ElseIf Not IsNumeric(params("dissertation_word_count")) Then
			showErrorPage "【论文字数】必须为数字（最多保留小数点后一位）！", "提示"
		End If

		tg.setField "StuNo", Session("StuNo")
		tg.setField "StuName", Session("StuName")
		If stu_type=5 Then
			tg.setField "StuTypeName", Format("{0}（{1}）", stu_type_name, params("research_field"))
		Else
			tg.setField "StuTypeName", stu_type_name
		End If
		tg.setField "TutorName", params("tutor_name")
		tg.setField "DefenceDate", Year(Now)&"年"&Month(Now)&"月  日"
		tg.setField "DegreeApplication", params("degree_application")
		tg.setField "Sex", params("sex")
		tg.setField "Ethnic", params("ethnic")
		tg.setField "Birthday", params("birthday")
		tg.setField "Nation", params("nation")
		tg.setField "PoliticalStatus", params("political_status")
		tg.setField "EntranceYearMonth", params("entrance_ym")
		tg.setField "IDCardNo", params("idcard_no")
		tg.setField "SpecialityName", params("speciality_name")
		tg.setField "StudyType", params("study_type")
		tg.setField "GraduatedAt", params("graduated_at")
		tg.setField "BeforeSpecialityName", params("before_speciality_name")
		tg.setField "LastDegree", params("last_degree")
		tg.setField "GraduationYearMonth", params("graduation_ym")

		tg.setArray "ResumeDuration", params("resume_duration")
		tg.setArray "ResumePlace", params("resume_place")
		tg.setArray "ResumeJob", params("resume_job")

		tg.setField "HonorPenalty", params("honor_penalty")

		tg.setArray "AchievementName", params("achievement_name")
		tg.setArray "AchievementYearMonth", params("achievement_ym")
		tg.setArray "AchievementSource", params("achievement_source")
		tg.setArray "AchievementAuthorNum", params("achievement_authornum")
		tg.setArray "AchievementStatus", params("achievement_status")

		tg.setField "DissertationSubject", params("dissertation_subject")
		tg.setField "DissertationKeywords", params("dissertation_keywords")
		tg.setField "DissertationWordCount", params("dissertation_word_count")
		tg.setField "IssueSource", params("issue_source")
		tg.setField "ProjectNameCode", params("project_name_code")
		tg.setField "DissertationType", params("dissertation_type")
		tg.setField "DissertationDuration", params("dissertation_duration")
		tg.setField "TutorEval", params("tutor_eval")

	End Select

	tg.generateTable export_file_path, template_name
	Set tg=Nothing

	If params.Exists("subject") Then subject=params("subject")
	If params.Exists("subject_en") Then subject_en=params("subject_en")

	Dim arrTableFieldName,arrNewTaskProgress
	arrTableFieldName=Array("","TABLE_FILE1","TABLE_FILE2","TABLE_FILE3","TABLE_FILE4")
	arrNewTaskProgress=Array(0,tpTbl1Uploaded,tpTbl2Uploaded,tpTbl3Uploaded,tpTbl4Uploaded)
	' 关联到数据库
	sql="SELECT * FROM Dissertations WHERE STU_ID="&Session("StuId")&" ORDER BY ActivityId DESC"
	GetRecordSet conn,rs3,sql,count
	If rs3.EOF Then
		' 添加记录
		rs3.AddNew()
	End If
	If is_new_dissertation Then
		rs3("STU_ID")=Session("StuId")
		rs3("ActivityId")=params("activity_id")
		rs3("REVIEW_STATUS")=rsNone
		rs3("REVIEW_RESULT")="5,5,6"
		rs3("REVIEW_LEVEL")="0,0"
		If params.Exists("keyword_ch_all") Then
			rs3("KEYWORDS")=Replace(params("keyword_ch_all"),", ","；")
			rs3("KEYWORDS_EN")=Replace(params("keyword_en_all"),", ","；")
		End If
	End If
	rs3("THESIS_SUBJECT")=subject
	rs3("THESIS_SUBJECT_EN")=subject_en
	If params("sub_research_field_select")="-1" Then
		rs3("RESEARCHWAY_NAME")=params("custom_sub_research_field")
	ElseIf params.Exists("sub_research_field") Then
		rs3("RESEARCHWAY_NAME")=params("sub_research_field")
	End If
	rs3(arrTableFieldName(section_id))=export_file_name
	rs3("TASK_PROGRESS")=arrNewTaskProgress(section_id)
	rs3.Update()
	CloseRs rs3

	' 保存视图状态
	setViewState Session("StuId"),usertypeStudent,view_name,params("view_state")
	writeLog Format("学生[{0}]填写提交[{1}]。",Session("Stuname"),arrStuOprName(section_id))
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>提交论文表格</title>
<% useStylesheet "student" %>
<% useScript "jquery" %>
</head>
<body><%
	If Not bError Then %>
<form id="fmFinish" action="home.asp" method="post">
<input type="hidden" name="filename" value="<%=export_file_name%>" />
</form>
<script type="text/javascript">alert("提交成功！");$('#fmFinish').submit();</script><%
	Else
%><script type="text/javascript">alert("<%=errMsg%>");history.go(-1);</script><%
	End If
%></body></html><%
End Select
CloseRs rs
CloseRs rsStu
CloseConn conn
%>
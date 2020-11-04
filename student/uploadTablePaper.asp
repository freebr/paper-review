<!--#include file="../inc/global.inc"-->
<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("StuId")) Then Response.Redirect("../error.asp?timeout")
Dim is_new_dissertation:is_new_dissertation=False
Dim activity_id,section_id,time_flag,uploadable
Dim conn,rs,sql,count

activity_id=0
section_id=0
uploadable=False
stu_type=Session("StuType")

Connect conn
sql="SELECT *,dbo.getThesisStatusText(1,TASK_PROGRESS,2) AS STAT_TEXT FROM ViewDissertations WHERE STU_ID="&Session("Stuid")&" ORDER BY ActivityId DESC"
GetRecordSetNoLock conn,rs,sql,count
If rs.EOF Then
	section_id=sectionUploadKtbgb
	task_progress=tpNone
Else
	activity_id=rs("ActivityId")
	subject_ch=rs("THESIS_SUBJECT")
	subject_en=rs("THESIS_SUBJECT_EN")
	' 表格审核进度
	task_progress=rs("TASK_PROGRESS")
	Select Case task_progress
	Case tpNone,tpTbl1Uploaded,tpTbl1Unpassed	' 开题报告
		section_id=sectionUploadKtbgb
	Case tpTbl1Passed,tpTbl2Uploaded,tpTbl2Unpassed	' 中期考核表
		section_id=sectionUploadZqkhb
	Case tpTbl2Passed,tpTbl3Uploaded,tpTbl3Unpassed	' 预答辩意见书
		section_id=sectionUploadYdbyjs
	End Select
End If
If section_id<>0 Then
	is_new_dissertation=section_id=sectionUploadKtbgb Or stu_type=7 And section_id=sectionUploadYdbyjs
	If rs.EOF Then
		uploadable=True
	ElseIf Not isActivityOpen(rs("ActivityId")) Then
		time_flag=-3
	Else
		Set current_section=getSectionInfo(rs("ActivityId"), stu_type, section_id)
		time_flag=compareNowWithSectionTime(current_section)
		uploadable=time_flag=0
	End If
End If
step=Request.QueryString("step")
Select Case step
Case vbNullstring ' 填写信息页面
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>上传表格附加论文</title>
<% useStylesheet "student", "jeasyui" %>
<% useScript "jquery", "jeasyui", "common", "upload", "uploadPaper" %>
</head>
<body>
<center><font size=4><b>上传表格附加论文</b></font>
<form id="fmDissertation" action="?step=1" method="post" enctype="multipart/form-data">
<table class="form" width="1000" align="center"><tr><td class="summary"><%
	If Not uploadable Then
%><p><span class="tip">当前状态为【<%=rs("STAT_TEXT")%>】，不能上传附加论文！</span></p><%
	ElseIf time_flag=-2 Then
%><p><span class="tip">【<%=current_section("Name")%>】环节已关闭，不能上传附加论文！</span></p><%
	ElseIf time_flag<>0 Then
%><p><span class="tip">【<%=current_section("Name")%>】环节开放时间为<%=toDateTime(current_section("StartTime"),1)%>至<%=toDateTime(current_section("EndTime"),1)%>，当前不在开放时间内，不能上传附加论文！</span></p><%
	Else
%><p>当前上传的是：<span style="color:#ff0000;font-weight:bold"><%=arrTblpaperDetail(section_id)%></span></p><%
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
<p>请选择要上传的文件，并点击&quot;提交&quot;按钮：</p><%
	End If %></td></tr>
<tr><td align="center">
<table class="form">
<tr><td><p>论文题目：《<input type="text" name="subject_ch" size="100" value="<%=subject_ch%>" />》</p>
<p>（英文）：&nbsp;<input type="text" name="subject_en" size="100" maxlength="200" value="<%=subject_en%>" />&nbsp;</p>
<p>文件名：<input type="file" name="thesis_file" size="50" title="<%=arrTblThesis(section_id)%>" /><br/><span class="tip">Word&nbsp;或&nbsp;RAR&nbsp;格式，超过20M请先压缩成rar文件再上传，否则上传不成功</span></p>
<p align="center"><input type="submit" id="btnsubmit" value="提 交"<%If Not uploadable Then %> disabled<% End If %> />&nbsp;
<input type="button" name="btnUploadTable" value="返回上传表格页面" onclick="location.href='uploadTable.asp'" />&nbsp;
<input type="button" name="btnreturn" value="返回首页" onclick="location.href='home.asp'" /></p></td></tr></table>
</td></tr></table></form></center>
<script type="text/javascript">
	$('input[name="thesis_file"]').change(function(){if(this.value.length)checkIfWordRar(this);});
	$('form').submit(function(){
			var valid=checkIfWordRar(this.thesis_file);
			if(valid) submitUploadForm(this); else return false;
		});
	<%
	If Not uploadable Then %>
	$('input[name="thesis_file"]').attr('readOnly',true);
	$(':submit').attr('disabled',true);<%
	Else %>
	$(':submit').attr('disabled',false);<%
	End If %>
</script></body></html><%
Case 1	' 上传进程

	If time_flag=-3 Then
		bError=True
		errMsg=Format("当前评阅活动【{0}】已关闭，不能提交表格！", rs("ActivityName"))
	ElseIf time_flag=-2 Then
		bError=True
		errMsg=Format("【{0}】环节已关闭，不能上传附加论文！",current_section("Name"))
	ElseIf time_flag<>0 Then
		bError=True
		errMsg=Format("【{0}】环节开放时间为{1}至{2}，当前不在开放时间内，不能上传附加论文！",_
			current_section("Name"),_
			toDateTime(current_section("StartTime"),1),_
			toDateTime(current_section("EndTime"),1))
	ElseIf Not uploadable Then
		bError=True
		errMsg="当前状态为【"&rs("STAT_TEXT")&"】，不能上传附加论文！"
	End If
	If bError Then
		CloseRs rs
		CloseConn conn
		showErrorPage errMsg, "提示"
	End If

	Dim fso,Upload,thesis_file
	Dim new_subject_ch,new_subject_en

	Set Upload=New ExtendedRequest
	activity_id=Upload.Form("activity_id")
	new_subject_ch=Upload.Form("subject_ch")
	new_subject_en=Upload.Form("subject_en")
	Set thesis_file=Upload.File("thesis_file")
	Set fso=CreateFSO()

	' 检查上传目录是否存在
	strUploadPath=Server.MapPath("upload")
	If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)
	thesis_file_ext=LCase(thesis_file.FileExt)
	If is_new_dissertation And Len(activity_id)=0 Then
		bError=True
		errMsg="请选择要参加的评阅活动！"
	ElseIf InStr("doc docx rar",thesis_file_ext)=0 Then	' 不被允许的文件类型
		bError=True
		errMsg="所上传的不是 Word 文件或 RAR 压缩文件！"
	ElseIf Len(new_subject_ch)=0 Then
		bError=True
		errMsg="请填写论文题目！"
	ElseIf Len(new_subject_en)=0 Then
		bError=True
		errMsg="请填写论文题目（英文）！"
'	ElseIf file.FileSize>10485760 Then
'		filesize=Round(file.FileSize/1048576,2)
'		bError=True
'		errMsg="文件大小为 "&filesize&"MB，已超出限制(10MB)！"
	Else
		byteFileSize=0
		' 生成日期格式文件名
		fileid=timestamp()
		strDestThesisFile=fileid&"."&thesis_file_ext
		destPath=strUploadPath&"\"&strDestThesisFile
		byteFileSize=thesis_file.FileSize
		' 保存论文文件
		thesis_file.SaveAs destPath
	End If
	Set fso=Nothing
	Set thesis_file=Nothing
	Set Upload=Nothing

	If Not bError Then
		Dim arrTblThesisFieldName
		arrTblThesisFieldName=Array("","TBL_THESIS_FILE1","TBL_THESIS_FILE2","TBL_THESIS_FILE3")
		' 关联到数据库
		sql="SELECT * FROM Dissertations WHERE STU_ID="&Session("Stuid")&" ORDER BY ActivityId DESC"
		GetRecordSet conn,rs3,sql,count
		If rs3.EOF Then
			' 添加记录
			rs3.AddNew()
		End If
		If is_new_dissertation Then	' 新论文记录，录入论文基本信息
			rs3("STU_ID")=Session("Stuid")
			rs3("ActivityId")=activity_id
			rs3("REVIEW_STATUS")=rsNone
			rs3("REVIEW_RESULT")="5,5,6"
			rs3("REVIEW_LEVEL")="0,0"
			rs3("RESEARCHWAY_NAME")=""
		End If
		rs3("THESIS_SUBJECT")=new_subject_ch
		rs3("THESIS_SUBJECT_EN")=new_subject_en
		rs3(arrTblThesisFieldName(section_id))=strDestThesisFile
		rs3.Update()
		CloseRs rs3

		writeLog Format("学生[{0}]上传[{1}]。",Session("Stuname"),arrTblThesis(section_id))
	End If
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>上传表格附加论文</title>
<% useStylesheet "student" %>
<% useScript "jquery" %>
</head>
<body><%
	If Not bError Then %>
<form id="fmFinish" action="home.asp" method="post">
<input type="hidden" name="filename" value="<%=strDestTableFile%>" />
</form>
<script type="text/javascript">alert("上传成功！");$('#fmFinish').submit();</script><%
	Else
%><script type="text/javascript">alert("<%=errMsg%>");history.go(-1);</script><%
	End If
%></body></html><%
End Select
CloseRs rs
CloseConn conn
%>
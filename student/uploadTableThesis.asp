<!--#include file="../inc/global.inc"-->
<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("StuId")) Then Response.Redirect("../error.asp?timeout")
Dim section_id,time_flag,allow_upload
Dim conn,rs,sql,count

allow_upload=False
section_id=0
stu_type=Session("StuType")

Connect conn
sql="SELECT *,dbo.getThesisStatusText(1,TASK_PROGRESS,2) AS STAT_TEXT FROM ViewThesisInfo WHERE STU_ID="&Session("Stuid")&" ORDER BY ActivityId DESC"
GetRecordSetNoLock conn,rs,sql,count
sql="SELECT TUTOR_ID FROM ViewStudentInfo WHERE STU_ID="&Session("Stuid")
If rs.EOF Then
	section_id=sectionUploadKtbg
	task_progress=tpNone
Else
	subject_ch=rs("THESIS_SUBJECT")
	subject_en=rs("THESIS_SUBJECT_EN")
	' 表格审核进度
	task_progress=rs("TASK_PROGRESS")
	Select Case task_progress
	Case tpNone,tpTbl1Uploaded,tpTbl1Unpassed	' 开题报告
		section_id=sectionUploadKtbg
	Case tpTbl1Passed,tpTbl2Uploaded,tpTbl2Unpassed	' 中期检查表
		section_id=sectionUploadZqjcb
	Case tpTbl2Passed,tpTbl3Uploaded,tpTbl3Unpassed	' 预答辩申请表
		section_id=sectionUploadYdbyjs
	Case Else
		allow_upload=False
	End Select
End If
If section_id<>0 Then
	If Not isActivityOpen(rs("ActivityId")) Then
		time_flag=-3
	Else
		Set current_section=getSectionInfo(rs("ActivityId"), stu_type, section_id)
		time_flag=compareNowWithSectionTime(current_section)
		allow_upload=time_flag=0
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
<% useStylesheet "student" %>
<% useScript "jquery", "common", "upload", "uploadThesis" %>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>上传表格附加论文</b></font>
<table class="tblform" width="1000" align="center"><tr><td class="summary"><p><%
	If Not allow_upload Then
%><span class="tip">当前状态为【<%=rs("STAT_TEXT")%>】，不能上传附加论文！</span><%
	ElseIf time_flag=-2 Then
%><span class="tip">【<%=current_section("Name")%>】环节已关闭，不能上传附加论文！</span><%
	ElseIf time_flag<>0 Then
%><span class="tip">【<%=current_section("Name")%>】环节开放时间为<%=toDateTime(current_section("StartTime"),1)%>至<%=toDateTime(current_section("EndTime"),1)%>，当前不在开放时间内，不能上传附加论文！</span><%
	Else
%>当前上传的是：<span style="color:#ff0000;font-weight:bold"><%=arrTblThesisDetail(section_id)%></span><br/>
请选择要上传的文件，并点击&quot;提交&quot;按钮：<%
	End If %></p></td></tr>
<tr><td align="center"><form id="fmDissertation" action="?step=1" method="post" enctype="multipart/form-data">
<table class="tblform">
<tr><td><p>论文题目：《<input type="text" name="subject_ch" size="50" value="<%=subject_ch%>" />》</p>
<p>（英文）：&nbsp;<input type="text" name="subject_en" size="53" maxlength="200" value="<%=subject_en%>" />&nbsp;</p>
<p>文件名：<input type="file" name="thesisFile" size="50" title="<%=arrTblThesis(section_id)%>" /><br/><span class="tip">Word&nbsp;或&nbsp;RAR&nbsp;格式，超过20M请先压缩成rar文件再上传，否则上传不成功</span></p>
<p align="center"><input type="submit" name="btnsubmit" value="提 交"<%If Not allow_upload Then %> disabled<% End If %> />&nbsp;
<input type="button" name="btnUploadTable" value="返回填写表格页面" onclick="location.href='uploadTableNew.asp'" />&nbsp;
<input type="button" name="btnreturn" value="返回首页" onclick="location.href='home.asp'" /></p></td></tr></table>
</form></td></tr></table></center>
<script type="text/javascript">
	$('input[name="thesisFile"]').change(function(){if(this.value.length)checkIfWordRar(this);});
	$('form').submit(function(){
			var valid=checkIfWordRar(this.thesisFile);
			if(valid) submitUploadForm(this); else return false;
		});
	<%
	If Not allow_upload Then %>
	$('input[name="thesisFile"]').attr('readOnly',true);
	$(':submit').attr('disabled',true);<%
	Else %>
	$(':submit').attr('disabled',false);<%
	End If %>
</script></body></html><%
Case 1	' 上传进程

	If time_flag=-3 Then
		bError=True
		errdesc=Format("当前评阅活动【{0}】已关闭，不能提交表格！", rs("ActivityName"))
	ElseIf time_flag=-2 Then
		bError=True
		errdesc=Format("【{0}】环节已关闭，不能上传附加论文！",current_section("Name"))
	ElseIf time_flag<>0 Then
		bError=True
		errdesc=Format("【{0}】环节开放时间为{1}至{2}，当前不在开放时间内，不能上传附加论文！",_
			current_section("Name"),_
			toDateTime(current_section("StartTime"),1),_
			toDateTime(current_section("EndTime"),1))
	ElseIf Not allow_upload Then
		bError=True
		errdesc="当前状态为【"&rs("STAT_TEXT")&"】，不能上传附加论文！"
	End If
	If bError Then
		CloseRs rs
		CloseConn conn
		showErrorPage errdesc, "提示"
	End If

	Dim fso,Upload,thesisfile
	Dim new_subject_ch,new_subject_en

	Set Upload=New ExtendedRequest
	new_subject_ch=Upload.Form("subject_ch")
	new_subject_en=Upload.Form("subject_en")
	Set thesisfile=Upload.File("thesisFile")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")

	' 检查上传目录是否存在
	strUploadPath=Server.MapPath("upload")
	If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)
	thesisfileExt=LCase(thesisfile.FileExt)
	If InStr("doc docx rar",thesisfileExt)=0 Then	' 不被允许的文件类型
		bError=True
		errdesc="上传论文所选择的不是 Word 文件或 RAR 压缩文件！"
	ElseIf Len(new_subject_ch)=0 Then
		bError=True
		errdesc="请填写论文题目！"
	ElseIf Len(new_subject_en)=0 Then
		bError=True
		errdesc="请填写论文题目（英文）！"
'	ElseIf file.FileSize>10485760 Then
'		filesize=Round(file.FileSize/1048576,2)
'		bError=True
'		errdesc="文件大小为 "&filesize&"MB，已超出限制(10MB)！"
	Else
		byteFileSize=0
		' 生成日期格式文件名
		fileid=FormatDateTime(Now(),1)&Int(Timer)
		strDestThesisFile=fileid&"."&thesisfileExt
		strDestPath=strUploadPath&"\"&strDestThesisFile
		byteFileSize=thesisfile.FileSize
		' 保存论文文件
		thesisfile.SaveAs strDestPath
	End If
	Set fso=Nothing
	Set thesisfile=Nothing
	Set Upload=Nothing

	If Not bError Then
		Dim arrTblThesisFieldName,arrNewTaskProgress
		arrTblThesisFieldName=Array("","TBL_THESIS_FILE1","TBL_THESIS_FILE2","TBL_THESIS_FILE3")
		arrNewTaskProgress=Array(0,tpTbl1Uploaded,tpTbl2Uploaded,tpTbl3Uploaded)
		' 关联到数据库
		sql="SELECT * FROM Dissertations WHERE STU_ID="&Session("Stuid")&" ORDER BY ActivityId DESC"
		GetRecordSet conn,rs3,sql,count
		If rs3.EOF Then
			' 添加记录
			rs3.AddNew()
		End If
		If section_id=sectionUploadKtbg Then	' 开题报告，录入论文基本信息
			rs3("STU_ID")=Session("Stuid")
			rs3("REVIEW_STATUS")=rsNone
			rs3("REVIEW_FILE_STATUS")=0
			rs3("REVIEW_RESULT")="5,5,6"
			rs3("REVIEW_LEVEL")="0,0"
			rs3("RESEARCHWAY_NAME")=""
		End If
		rs3("THESIS_SUBJECT")=new_subject_ch
		rs3("THESIS_SUBJECT_EN")=new_subject_en
		rs3(arrTblThesisFieldName(section_id))=strDestThesisFile
		'rs3("TASK_PROGRESS")=arrNewTaskProgress(section_id)
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
<body bgcolor="ghostwhite"><%
	If Not bError Then %>
<form id="fmFinish" action="home.asp" method="post">
<input type="hidden" name="filename" value="<%=strDestTableFile%>" />
</form>
<script type="text/javascript">alert("上传成功！");$('#fmFinish').submit();</script><%
	Else
%><script type="text/javascript">alert("<%=errdesc%>");history.go(-1);</script><%
	End If
%></body></html><%
End Select
CloseRs rs
CloseConn conn
%>
<%Response.Charset="utf-8"%>
<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("StuId")) Then Response.Redirect("../error.asp?timeout")
Dim opr,bOpen,bUpload
Dim conn,rs,sql,result
bOpen=True
bUpload=True
opr=0
sem_info=getCurrentSemester()
stu_type=Session("StuType")
Connect conn
sql="SELECT *,dbo.getThesisStatusText(1,TASK_PROGRESS,2) AS STAT_TEXT FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE STU_ID="&Session("Stuid")&" ORDER BY PERIOD_ID DESC" 'AND PERIOD_ID="&sem_info(3)&" AND Valid=1"
GetRecordSetNoLock conn,rs,sql,result
sql="SELECT * FROM VIEW_STUDENT_INFO WHERE STU_ID="&Session("Stuid")
GetRecordSetNoLock conn,rs2,sql,result
tutor_duty_name=getProDutyNameOf(rs2("TUTOR_ID"))
If rs.EOF Then
	opr=STUCLI_OPR_TABLE1
	task_progress=tpNone
Else
	subject_ch=rs("THESIS_SUBJECT")
	subject_en=rs("THESIS_SUBJECT_EN")
	' 表格审核进度
	task_progress=rs("TASK_PROGRESS")
	Select Case task_progress
	Case tpNone,tpTbl1Uploaded,tpTbl1Unpassed	' 开题报告
		opr=STUCLI_OPR_TABLE1
	Case tpTbl1Passed,tpTbl2Uploaded,tpTbl2Unpassed	' 中期检查表
		opr=STUCLI_OPR_TABLE2
	Case tpTbl2Passed,tpTbl3Uploaded,tpTbl3Unpassed	' 预答辩申请表
		opr=STUCLI_OPR_TABLE3
	Case Else
		bUpload=False
	End Select
End If
If opr<>0 Then
	bOpen=stuclient.isOpenFor(stu_type,opr)
	startdate=stuclient.getOpentime(opr,STUCLI_OPENTIME_START)
	enddate=stuclient.getOpentime(opr,STUCLI_OPENTIME_END)
	If Not bOpen Then bUpload=False
End If
curStep=Request.QueryString("step")
Select Case curStep
Case vbNullstring ' 填写信息页面
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/student.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.11.3.min.js" type="text/javascript"></script>
<script src="../scripts/upload.js" type="text/javascript"></script>
<script src="../scripts/uploadThesis.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>上传表格附加论文</b></font>
<table class="tblform" width="1000" align="center"><tr><td class="summary"><p><%
	If Not bOpen Then
%><span class="tip">上传<%=arrTblThesis(opr)%>的时间为<%=toDateTime(startdate,1)%>至<%=toDateTime(enddate,1)%>，本专业上传通道已关闭或当前不在开放时间内，不能上传论文！</span><%
	ElseIf Not bUpload Then
%><span class="tip">当前状态为【<%=rs("STAT_TEXT")%>】，不能上传论文！</span><%
	Else
%>当前上传的是：<span style="color:#ff0000;font-weight:bold"><%=arrTblThesisDetail(opr)%></span><br/>
请选择要上传的文件，并点击&quot;提交&quot;按钮：<%
	End If %></p></td></tr>
<tr><td align="center"><form id="fmThesis" action="?step=1" method="post" enctype="multipart/form-data">
<table class="tblform">
<tr><td><p>论文题目：《<input type="text" name="subject_ch" size="50" value="<%=subject_ch%>" />》</p>
<p>（英文）：&nbsp;<input type="text" name="subject_en" size="53" maxlength="200" value="<%=subject_en%>" />&nbsp;</p>
<p>文件名：<input type="file" name="thesisFile" size="50" title="<%=arrTblThesis(opr)%>" /><br/><span class="tip">Word&nbsp;或&nbsp;RAR&nbsp;格式，超过20M请先压缩成rar文件再上传，否则上传不成功</span></p>
<p align="center"><input type="submit" name="btnsubmit" value="提 交"<%If Not bUpload Then %> disabled<% End If %> />&nbsp;
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
	If Not bUpload Then %>
	$('input[name="thesisFile"]').attr('readOnly',true);
	$(':submit').attr('disabled',true);<%
	Else %>
	$(':submit').attr('disabled',false);<%
	End If %>
</script></body></html><%
Case 1	' 上传进程
	If Not bOpen Then
		bError=True
		errdesc="上传"&arrTblThesis(opr)&"的时间为"&FormatDateTime(startdate,1)&"至"&FormatDateTime(enddate,1)&"，本专业上传通道已关闭或当前不在开放时间内，不能上传论文！"
	ElseIf Not bUpload Then
		bError=True
		errdesc="当前状态为【"&rs("STAT_TEXT")&"】，不能上传论文！"
	End If
	If bError Then
		%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
  	CloseRs rs
  	CloseRs rs2
  	CloseConn conn
		Response.End()
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
	
	Dim arrTblThesisFieldName,arrNewTaskProgress
	arrTblThesisFieldName=Array("","TBL_THESIS_FILE1","TBL_THESIS_FILE2","TBL_THESIS_FILE3")
	arrNewTaskProgress=Array(0,tpTbl1Uploaded,tpTbl2Uploaded,tpTbl3Uploaded)
	' 关联到数据库
	sql="SELECT * FROM TEST_THESIS_REVIEW_INFO WHERE STU_ID="&Session("Stuid")&" ORDER BY PERIOD_ID DESC"
	GetRecordSet conn,rs3,sql,result
	If rs3.EOF Then
		' 添加记录
		rs3.AddNew()
	End If
	If opr=STUCLI_OPR_TABLE1 Then	' 开题报告，录入论文基本信息
		rs3("STU_ID")=Session("Stuid")
		rs3("REVIEW_STATUS")=rsNone
		rs3("REVIEW_FILE_STATUS")=0
		rs3("REVIEW_RESULT")="5,5,6"
		rs3("REVIEW_LEVEL")="0,0"
		rs3("RESEARCHWAY_NAME")=""
	End If
	rs3("PERIOD_ID")=sem_info(3)
	rs3("THESIS_SUBJECT")=new_subject_ch
	rs3("THESIS_SUBJECT_EN")=new_subject_en
	rs3(arrTblThesisFieldName(opr))=strDestThesisFile
	'rs3("TASK_PROGRESS")=arrNewTaskProgress(opr)
	rs3.Update()
	CloseRs rs3
	' 向导师发送审核通知邮件
	'sendEmailToTutor arrTblThesis(opr)
	Dim logtxt
	logtxt="学生["&Session("Stuname")&"]上传["&arrTblThesis(opr)&"]。"
	writeLog logtxt
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>上传表格附加论文</title>
<link href="../css/student.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.11.3.min.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite"><%
	If Not bError Then %>
<form id="fmFinish" action="home.asp" method="post">
<input type="hidden" name="filename" value="<%=strDestTableFile%>" />
<p><%=byteFileSize%> 字节已上传，正在关联数据...</p></form>
<script type="text/javascript">alert("上传成功！");$('#fmFinish').submit();</script><%
	Else
%><script type="text/javascript">alert("<%=errdesc%>");history.go(-1);</script><%
	End If
%></body></html><%
End Select
CloseRs rs2
CloseRs rs
CloseConn conn
%>
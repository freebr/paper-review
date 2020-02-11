<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")

Dim conn,sql,ret,rs

step=Request.QueryString("step")
Select Case step
Case vbNullstring ' 文件选择页面
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入答辩委员会修改意见</title>
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>导入答辩委员会修改意见</b><br />
<form id="fmUpload" action="?step=2" method="POST" enctype="multipart/form-data">
<p>检索方式：<select name="selectmode"><option value="0" selected>按学号检索</option><option value="1">按姓名检索</option></select></p>
<p><label for="chksendemail"><input type="checkbox" name="sendemail" id="chksendemail" checked />导入后发送通知邮件给导师和学生</label></p>
<p>请选择要导入的 Excel 文件：<br />文件名：<input type="file" name="excelFile" size="100" /><br />
<a href="upload/defenceeval_template.xlsx" target="_blank">点击下载答辩委员会修改意见表格模板</a><br />
<input type="submit" name="btnsubmit" value="提 交" />&nbsp;
<input type="button" name="btnret" value="返 回" onclick="history.go(-1)" /></p></form></center>
<script type="text/javascript">
	$(document).ready(function(){
		$('form').submit(function() {
			var valid=checkIfExcel(this.excelFile);
			if(valid) {
				$(':submit').val("正在提交，请稍候...").attr('disabled',true);
			}
			return valid;
		});
		$(':submit').attr('disabled',false);
	});
</script></body></html><%
Case 2	' 上传进程

	Dim fso,Upload,file

	Set Upload=New ExtendedRequest
	Set file=Upload.File("excelFile")
	select_mode=Upload.Form("selectmode")
	send_email=Upload.Form("sendemail")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")

	' 检查上传目录是否存在
	strUploadPath = Server.MapPath("upload\xls")
	If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)

	fileExt=LCase(file.FileExt)
	If fileExt <> "xls" And fileExt <> "xlsx" Then	' 不被允许的文件类型
		bError = True
		errstring = "所选择的不是 Excel 文件！"
	Else
		' 生成日期格式文件名
		fileid = FormatDateTime(Now(),1)&Int(Timer)
		strDestFile = fileid&"."&fileExt
		strDestPath = Server.MapPath("upload")&"\xls\"&strDestFile
		byteFileSize = file.FileSize
		' 保存
		file.SaveAs strDestPath
	End If
	Set file=Nothing
	Set Upload=Nothing
	Set fso=Nothing
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入答辩委员会修改意见</title>
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body bgcolor="ghostwhite">
<center><br /><b>导入答辩委员会修改意见</b><br /><br /><%
	If Not bError Then %>
<form id="fmUploadFinish" action="?step=3" method="POST">
<input type="hidden" name="selectmode" value="<%=select_mode%>" />
<input type="hidden" name="sendemail" value="<%=send_email%>" />
<input type="hidden" name="filename" value="<%=strDestFile%>" />
<p>文件上传成功，正在导入答辩委员会修改意见...</p></form>
<script type="text/javascript">setTimeout("$('#fmUploadFinish').submit()",500);</script><%
	Else
%>
<script type="text/javascript">alert("<%=errstring%>");history.go(-1);</script><%
	End If
%></center></body></html><%
Case 3	' 数据读取，导入到数据库

	Dim bError,errMsg
	Dim countInsert,countUpdate,thesis_ids
	
	filename=Request.Form("filename")
	filepath=Server.MapPath("upload/xls/"&filename)
	select_mode=Request.Form("selectmode")
	send_email=Request.Form("sendemail")="on"
	sql="CREATE TABLE #ret(CountInsert int,CountUpdate int,CountError int,FirstImportThesisIDs nvarchar(MAX),IsError bit,ErrMsg nvarchar(MAX));"&_
			"INSERT INTO #ret EXEC spImportDefenceEval '"&filepath&"',"&select_mode&"; SELECT * FROM #ret"
	Connect conn
	Set rs=conn.Execute(sql).NextRecordSet()
	countInsert=rs("CountInsert")
	countUpdate=rs("CountUpdate")
	thesis_ids=rs("FirstImportThesisIDs")
	bError=rs("IsError")
	errMsg=rs("ErrMsg")
	CloseRs rs

	If send_email And Len(thesis_ids) Then
		' 发送导入答辩意见通知邮件
		Dim arrDissertations:arrDissertations=Split(thesis_ids,",")
		Dim dict:Set dict=CreateDictionary()
		Dim operation_name,activity_id,stu_type,is_sent
		dict("filename")="答辩委员会修改意见"
		operation_name=Format("导入[{0}]",dict("filename"))
		sql="SELECT ActivityId,TEACHTYPE_ID,STU_NAME,STU_NO,CLASS_NAME,SPECIALITY_NAME,EMAIL,THESIS_SUBJECT,TUTOR_NAME,TUTOR_EMAIL FROM ViewDissertations WHERE ID=?"
		For i=0 To UBound(arrDissertations)
			Set ret=ExecQuery(conn,sql,CmdParam("ID",adInteger,4,arrDissertations(i)))
			Set rs=ret("rs")
			If Not rs.EOF Then
				activity_id=rs("ActivityId")
				stu_type=rs("TEACHTYPE_ID")
				dict("stuname")=rs("STU_NAME")
				dict("stuno")=rs("STU_NO")
				dict("stuclass")=rs("CLASS_NAME")
				dict("stuspec")=rs("SPECIALITY_NAME")
				dict("stumail")=rs("EMAIL")
				dict("subject")=rs("THESIS_SUBJECT")
				dict("tutorname")=rs("TUTOR_NAME")
				dict("tutormail")=rs("TUTOR_EMAIL")
				CloseRs rs

				is_sent=sendNotifyMail(activity_id,stu_type,"xxdrtzyj(xs)",dict("stumail"),dict)
				writeNotificationEventLog usertypeAdmin,Session("name"),operation_name,usertypeStudent,_
					dict("stuname"),dict("stumail"),notifytypeMail,is_sent

				is_sent=sendNotifyMail(activity_id,stu_type,"xxdrtzyj(ds)",dict("tutormail"),dict)
				writeNotificationEventLog usertypeAdmin,Session("name"),operation_name,usertypeTutor,_
					dict("tutorname"),dict("tutormail"),notifytypeMail,is_sent
			End If
		Next
	End If
	CloseConn conn
%><script type="text/javascript"><%
	If bError Then %>
	alert("导入时出错，其余 <%=countInsert%> 条记录已导入，<%=countUpdate%> 条记录已更新。以下是错误详情：\n<%=toJsString(errMsg)%>");
<%Else %>
	alert("操作成功，<%=countInsert%> 条记录已导入，<%=countUpdate%> 条记录已更新。");
<%End If
%>location.href="thesisList.asp";
</script><%
End Select
%>
<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")

tableUploadPath = Server.MapPath(uploadBasePath(usertypeAdmin, "defence_plan"))
ensurePathExists tableUploadPath

step=Request.QueryString("step")
Select Case step
Case vbNullstring ' 文件选择页面
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入答辩安排信息</title>
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body>
<center><font size=4><b>导入答辩安排信息</b><br />
<form id="fmUpload" action="?step=2" method="POST" enctype="multipart/form-data">
<p><label for="doNoticeStudent"><input type="checkbox" name="do_notice_student" id="doNoticeStudent" checked />导入后发送通知邮件给导师和学生</label></p>
<p>请选择要导入的 Excel 文件：<input type="file" name="tableFile" size="100" /></p>
<p><a href="upload/defenceplan_template.xlsx" target="_blank">点击下载答辩安排信息表格模板</a></p>
<p><input type="submit" name="btnsubmit" value="提 交" />&nbsp;
<input type="button" name="btnret" value="返 回" onclick="history.go(-1)" /></p></form></center>
<script type="text/javascript">
	$(document).ready(function(){
		$('form').submit(function() {
			var valid=checkIfExcel(this.tableFile);
			if(valid) {
				$(':submit').val("正在提交，请稍候...").attr('disabled',true);
			}
			return valid;
		});
		$(':submit').attr('disabled',false);
	});
</script></body></html><%
Case 2	' 上传进程

	Dim Upload,File

	Set Upload=New ExtendedRequest
	Set file=Upload.File("tableFile")
	do_notice_student=Upload.Form("do_notice_student")

	file_ext=LCase(file.FileExt)
	If file_ext <> "xls" And file_ext <> "xlsx" Then	' 不被允许的文件类型
		bError = True
		errMsg = "所选择的不是 Excel 文件！"
	Else
		destFile = timestamp()&"."&file_ext
		destPath = resolvePath(tableUploadPath,destFile)
		file.SaveAs destPath
	End If
	Set file=Nothing
	Set Upload=Nothing
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入答辩安排信息</title>
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body>
<center><br /><b>导入答辩安排信息</b><br /><br /><%
	If Not bError Then %>
<form id="fmUploadFinish" action="?step=3" method="POST">
<input type="hidden" name="do_notice_student" value="<%=do_notice_student%>" />
<input type="hidden" name="filename" value="<%=destFile%>" />
<p>文件上传成功，正在导入答辩安排信息...</p></form>
<script type="text/javascript">setTimeout("$('#fmUploadFinish').submit()",500);</script><%
	Else
%>
<script type="text/javascript">alert("<%=errMsg%>");history.go(-1);</script><%
	End If
%></center></body></html><%
Case 3	' 数据读取，导入到数据库

	Function transchar(ByVal s)
		' 将 | 号替换为实体符号
		s=Replace(s,"|","&brvbar;")
		transchar=s
	End Function

	Function addData()
		' 添加数据
		Dim sql,sql2,count,rsa,rsb
		Dim numInsert,numUpdate:numInsert=0:numUpdate=0
		Dim thesisid,stu_id,stu_ids:numPapers=0
		Dim member_desc,last_val(7)
		Do While Not rs.EOF
			If IsNull(rs(0)) Then Exit Do
			' 按学号检索
			sql="SELECT ID,STU_ID FROM ViewDissertations WHERE STU_NO="&toSqlString(rs(2))&" AND TEACHTYPE_ID="&getTeachTypeIdByName(rs(0))&" ORDER BY STU_ID DESC"
			GetRecordSetNoLock conn,rsa,sql,count
			If rsa.EOF Then
				bError=True
				errMsg=errMsg&"学生不存在：["&rs(1)&"]。"&vbNewLine
			Else
				numPapers=numPapers+1
				thesisid=rsa("ID")
				stu_id=rsa("STU_ID")
				stu_ids=stu_ids&","&stu_id
				sql="SELECT * FROM DefenceInfo WHERE THESIS_ID="&thesisid
				GetRecordSet conn,rsb,sql,count
				If rsb.EOF Then
					rsb.AddNew()
					numInsert=numInsert+1
				Else
					numUpdate=numUpdate+1
				End If

				member_desc=transchar(rs(4))&"|"&transchar(rs(5))&"|"&transchar(rs(6))
				rsb("THESIS_ID")=thesisid
				rsb("DEFENCE_TIME")=rs(7)
				rsb("DEFENCE_PLACE")=rs(8)
				rsb("DEFENCE_MEMBER")=member_desc
				rsb("MEMO")=rs(9)
				rsb.Update()
				CloseRs rsb
			End If
			CloseRs rsa
			rs.MoveNext()
		Loop
		addData=Array(numInsert,numUpdate,Mid(stu_ids,2))
	End Function

	Dim bError,errMsg
	Dim conn,connExcel,sql,ret,rs
	Dim numInsert,numUpdate

	filename=Request.Form("filename")
	do_notice_student=Request.Form("sendemail")="on"
	filepath=resolvePath(tableUploadPath,filename)
	Set connExcel=Server.CreateObject("ADODB.Connection")
	connstring="Provider=Microsoft.ACE.OLEDB.12.0;Data Source="&filepath&";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"""
	connExcel.Open connstring

	Set rs=connExcel.OpenSchema(adSchemaTables)
	Do While Not rs.EOF
		If rs("TABLE_TYPE")="TABLE" Then
			table_name=rs("TABLE_NAME")
			If InStr("Sheet1$",table_name) Then Exit Do
		End If
		rs.MoveNext()
	Loop
	sql="SELECT * FROM ["&table_name&"A2:J]"
	Set rs=connExcel.Execute(sql)
	' 添加数据
	ConnectDb conn
	ret=addData()
	numInsert=ret(0)
	numUpdate=ret(1)
	stu_ids=ret(2)

	If do_notice_student And Len(stu_ids)<>0 Then
		' 发送导入答辩安排通知邮件
		Dim arrStuIds:arrStuIds=Split(stu_ids,",")
		Dim dict:Set dict=CreateDictionary()
		Dim operation_name,activity_id,stu_type,is_sent
		dict("filename")="答辩安排信息"
		operation_name=Format("导入[{0}]",dict("filename"))
		sql="SELECT ActivityId,TEACHTYPE_ID,STU_NAME,STU_NO,CLASS_NAME,SPECIALITY_NAME,EMAIL,THESIS_SUBJECT,TUTOR_NAME,TUTOR_EMAIL FROM ViewDissertations WHERE STU_ID=?"
		For i=0 To UBound(arrStuIds)
			Set ret=ExecQuery(conn,sql,CmdParam("ID",adInteger,4,arrStuIds(i)))
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

	CloseRs rs
	CloseConn conn
	CloseConn connExcel
%><script type="text/javascript"><%
	If bError Then %>
	alert("导入时出错，其余<%=numInsert%>条记录已导入成功，<%=numUpdate%>条记录已更新。出错原因为：\n<%=toJsString(errMsg)%>");
<%Else %>
	alert("操作成功，<%=numInsert%>条记录已导入，<%=numUpdate%>条记录已更新。");
<%End If
%>location.href="paperList.asp";
</script><%
End Select
%>
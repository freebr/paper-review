<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%

step=Request.QueryString("step")
Select Case step
Case vbNullstring ' 文件选择页面
	activity_id=toUnsignedInt(Request.Form("In_ActivityId2"))
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入新增论文信息</title>
<% useStylesheet "admin" %>
<% useScript "jquery", "upload" %>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>导入新增论文信息</b><br />
<form id="fmUpload" action="?step=2" method="POST" enctype="multipart/form-data">
<p>评阅活动：<%=activityList("In_ActivityId", Session("AdminType")("ManageStuTypes"), activity_id, False)%></p>
<p>表格审核状态：<select name="In_TASK_PROGRESS"><%
GetMenuListPubTerm "ReviewStatuses","STATUS_ID1","STATUS_NAME","","AND STATUS_ID1 IS NOT NULL"
%></select></p>
<p>论文审核状态：<select name="In_REVIEW_STATUS"><%
GetMenuListPubTerm "ReviewStatuses","STATUS_ID2","STATUS_NAME","","AND STATUS_ID2 IS NOT NULL"
%></select></p>
<p>检索方式：<select name="selectmode"><option value="0" selected>按学号检索</option><option value="1">按姓名检索</option></select></p>
<p>请选择要导入的 Excel 文件：<br />文件名：<input type="file" name="excelFile" size="100" /><br />
<a href="upload/newpaperinf_template.xlsx" target="_blank">点击下载论文信息表格模板</a><br />
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

	Dim fso,Upload,File

	Set Upload=New ExtendedRequest
	Set file=Upload.File("excelFile")
	activity_id=toUnsignedInt(Upload.Form("In_ActivityId"))
	task_progress=Upload.Form("In_TASK_PROGRESS")
	review_status=Upload.Form("In_REVIEW_STATUS")
	select_mode=Upload.Form("selectmode")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")

	' 检查上传目录是否存在
	strUploadPath = Server.MapPath("upload\xls")
	If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)

	fileExt=LCase(file.FileExt)
	If activity_id="0" Then
		bError = True
		errstring = "请选择评阅活动！"
	ElseIf fileExt <> "xls" And fileExt <> "xlsx" Then	' 不被允许的文件类型
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
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入新增论文信息</title>
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body bgcolor="ghostwhite">
<center><br /><b>导入新增论文信息</b><br /><br /><%
	If Not bError Then %>
<form id="fmUploadFinish" action="?step=3" method="POST">
<input type="hidden" name="In_ActivityId" value="<%=activity_id%>" />
<input type="hidden" name="In_TASK_PROGRESS" value="<%=task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS" value="<%=review_status%>" />
<input type="hidden" name="selectmode" value="<%=select_mode%>" />
<input type="hidden" name="filename" value="<%=strDestFile%>" />
<p>文件上传成功，正在导入新增论文信息...</p></form>
<script type="text/javascript">setTimeout("$('#fmUploadFinish').submit()",500);</script><%
	Else
%>
<script type="text/javascript">alert("<%=errstring%>");history.go(-1);</script><%
	End If
%></center></body></html><%
Case 3	' 数据读取，导入到数据库

	Function addData()
		' 添加数据
		Dim fieldValue(3)
		Dim sql,sql_upd_rv,sql_upd_pv,sql_upd_apply,conn,connOrigin,count,rsa,rsb,rsc
		Dim stuid,tutorid,recid,teachtypeid,submit_review_time
		Dim numThesis
		Dim s,i,strTmp,strTmp2
		If review_status>=rsAgreedReview Then
			submit_review_time=Now
		Else
			submit_review_time=vbNullString
		End If
		submit_review_time=toSqlString(submit_review_time)
		numThesis=0
		sql_upd_rv="DECLARE @id int;"
		Connect conn
		ConnectOriginDb connOrigin
		Do While Not rs.EOF
			If IsNull(rs(0)) Then Exit Do
			' 学号
			s=toSqlString(Trim(rs(1)))
			If select_mode=0 Then ' 按学号检索
				sql="SELECT STU_ID,WRITEPRIVILEGETAGSTRING,READPRIVILEGETAGSTRING,dbo.getTeachTypeId(TEACHTYPE_ID,CLASS_NAME) AS TEACHTYPE_ID FROM ViewStudentInfo WHERE VALID=0 AND STU_NO="&s
			Else	' 按姓名检索（不可靠）
				sql="SELECT STU_ID,WRITEPRIVILEGETAGSTRING,READPRIVILEGETAGSTRING,dbo.getTeachTypeId(TEACHTYPE_ID,CLASS_NAME) AS TEACHTYPE_ID FROM ViewStudentInfo WHERE STU_ID=(SELECT TOP 1 STU_ID FROM ViewStudentInfo WHERE VALID=0 AND STU_NAME="&toSqlString(rs(0))&" AND TEACHTYPE_ID="&getTeachTypeIdByName(rs(3))&" ORDER BY STU_ID DESC)"
			End If
			Set rsa=conn.Execute(sql)
			If rsa.EOF Then
				bError=True
				errMsg=errMsg&"学生不存在:"""&rs(0)&"""。"&vbNewLine
			Else
				stuid=rsa("STU_ID")
				' 导师姓名
				fieldValue(0)=toSqlString(rs(2))
				' 学位类别
				fieldValue(1)=rsa(3)
				' 论文形式
				If rs(4)="无" Or IsNull(rs(4)) Then
					fieldValue(2)="''"
				Else
					fieldValue(2)=toSqlString(rs(4))
				End If
				' 论文题目
				fieldValue(3)=toSqlString(rs(5))
				sql="SELECT TEACHER_ID FROM ViewTutorList WHERE TEACHER_NAME="&fieldValue(0)
				Set rsb=conn.Execute(sql)
				If Not rsb.EOF Then
					tutorid=rsb("TEACHER_ID")
					sql="SELECT RECRUIT_ID,TEACHTYPE_ID FROM TutorRecruitSys..ViewRecruitInfo WHERE TEACHER_ID="&tutorid&" AND PERIOD_ID="&activity("SemesterId")&" AND TEACHTYPE_ID="&fieldValue(1)
					Set rsc=conn.Execute(sql)
					If Not rsc.EOF Then
						recid=rsc("RECRUIT_ID")
						teachtypeid=rsc("TEACHTYPE_ID")

						sql_upd_rv=sql_upd_rv&"SET @id=NULL;SELECT @id=ID FROM Dissertations WHERE STU_ID="&stuid&"; IF @id IS NULL INSERT INTO Dissertations (STU_ID,THESIS_SUBJECT,REVIEW_TYPE,TASK_PROGRESS,REVIEW_STATUS,SUBMIT_REVIEW_TIME,REVIEW_FILE_STATUS,REVIEW_RESULT,REVIEW_LEVEL,ActivityId,VALID) VALUES("&_
						stuid&","&fieldValue(3)&",dbo.getReviewTypeId("&teachtypeid&","&fieldValue(2)&"),"&task_progress&","&review_status&","&submit_review_time&",0,'5,5,6','0,0',"&activity_id&",1);"&_
						"ELSE UPDATE Dissertations SET THESIS_SUBJECT="&fieldValue(3)&",REVIEW_TYPE=dbo.getReviewTypeId("&teachtypeid&","&fieldValue(2)&"),TASK_PROGRESS="&task_progress&",REVIEW_STATUS="&review_status&",SUBMIT_REVIEW_TIME=CASE WHEN SUBMIT_REVIEW_TIME IS NULL THEN "&submit_review_time&" ELSE SUBMIT_REVIEW_TIME END,ActivityId="&activity_id&",VALID=1 WHERE ID=@id;"

						sql_upd_pv=sql_upd_pv&"UPDATE STUDENT_INFO SET TUTOR_ID="&tutorid&",TUTOR_RECRUIT_ID="&recid&",TUTOR_RECRUIT_STATUS=3,"&_
											 "WRITEPRIVILEGETAGSTRING=dbo.addPrivilege(WRITEPRIVILEGETAGSTRING,'SA8',''),READPRIVILEGETAGSTRING=dbo.addPrivilege(READPRIVILEGETAGSTRING,'SA8','') WHERE STU_ID="&stuid&";"

						sql_upd_apply=sql_upd_apply&"IF NOT EXISTS(SELECT STU_ID FROM TutorRecruitSys..ApplyInfo WHERE STU_ID="&stuid&" AND RECRUIT_ID="&recid&") BEGIN;"&_
													"DELETE FROM TutorRecruitSys..ApplyInfo WHERE STU_ID="&stuid&" AND TURN_NUM=1;"&_
													"INSERT INTO TutorRecruitSys..ApplyInfo (STU_ID,TUTOR_ID,RECRUIT_ID,PERIOD_ID,TURN_NUM,APPLY_TIME,TUTOR_REPLY_TIME,APPLY_STATUS) VALUES("&stuid&","&tutorid&","&recid&","&activity("SemesterId")&",1,'"&Now&"','"&Now&"',3); END;"

						numThesis=numThesis+1
					Else
						bError=True
						errMsg=errMsg&"学生"""&rs(0)&"""所选导师"""&rs(2)&"""缺少必需的招生信息。"&vbNewLine
					End If
					CloseRs rsc
				Else
					bError=True
					errMsg=errMsg&"学生"""&rs(0)&"""所选导师"""&rs(2)&"""未被录入导师信息数据库。"&vbNewLine
				End If
				CloseRs rsb
			End If
			CloseRs rsa
			rs.MoveNext()
		Loop
		' 增加新的评阅论文，并更新已有评阅论文
		If Len(sql_upd_rv) Then conn.Execute sql_upd_rv
		' 添加学生访问评阅系统的权限
		If Len(sql_upd_pv) Then connOrigin.Execute sql_upd_pv
		' 添加学生选导师系统填报志愿信息
		If Len(sql_upd_apply) Then conn.Execute sql_upd_apply
		CloseConn connOrigin
		CloseConn conn
		addData=numThesis
	End Function

	Dim bError,errMsg

	filename=Request.Form("filename")
	activity_id=toUnsignedInt(Request.Form("In_ActivityId"))
	task_progress=Request.Form("In_TASK_PROGRESS")
	review_status=Request.Form("In_REVIEW_STATUS")
	select_mode=Request.Form("selectmode")
	filepath=Server.MapPath("upload/xls/"&filename)

	Set activity=getActivityInfo(activity_id)
	If IsNull(activity) Then
		showErrorPage "所选评阅活动无效！", "提示"
	End If

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
	sql="SELECT * FROM ["&table_name&"A2:F]"
	Set rs=connExcel.Execute(sql)
	' 添加数据
	ret=addData()
	CloseRs rs
	CloseConn connExcel
%><script type="text/javascript"><%
	If bError Then %>
	alert("导入时出错，其余<%=ret%>条记录已导入成功。出错原因为：\n<%=toJsString(errMsg)%>");
<%Else %>
	alert("操作成功，<%=ret%>条记录已导入。");
<%End If
%>location.href="paperList.asp";
</script><%
End Select
%>
<!--#include file="../inc/global.inc"-->
<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
paper_id=Request.QueryString("tid")
step=Request.QueryString("step")

Select Case step
Case vbNullString	' 论文详情页面
	activity_id=Request.Form("In_ActivityId2")
	teachtype_id=Request.Form("In_TEACHTYPE_ID2")
	class_id=Request.Form("In_CLASS_ID2")
	enter_year=Request.Form("In_ENTER_YEAR2")
	query_task_progress=Request.Form("In_TASK_PROGRESS2")
	query_review_status=Request.Form("In_REVIEW_STATUS2")
	finalFilter=Request.Form("finalFilter2")
	pageSize=Request.Form("pageSize2")
	pageNo=Request.Form("pageNo2")
	If Len(paper_id)=0 Or Not IsNumeric(paper_id) Then
	%><body><center><font color=red size="4">参数无效。</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		Response.End()
	End If
	Connect conn
	sql="SELECT *,dbo.getThesisStatusText(1,TASK_PROGRESS,1) AS STAT_TEXT1,dbo.getThesisStatusText(2,REVIEW_STATUS,1) AS STAT_TEXT2 FROM ViewDissertations WHERE ID="&paper_id
	GetRecordSet conn,rs,sql,count
	If count=0 Then
	%><body><center><font color=red size="4">数据库没有该论文记录！</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		CloseRs rs
		CloseConn conn
		Response.End()
	End If
	task_progress=rs("TASK_PROGRESS")
	review_status=rs("REVIEW_STATUS")
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>上传表格/论文文件</title>
<% useStylesheet "admin" %>
<% useScript "jquery", "common", "paper" %>
</head>
<body>
<center><font size=4><b>上传表格/论文文件</font>
<form id="fmDetail" action="?step=2&tid=<%=paper_id%>" enctype="multipart/form-data" method="post">
<table class="form" width="800" cellspacing="1" cellpadding="3">
<tr><td>论文题目：&emsp;&emsp;&emsp;<input type="text" class="txt full-width" name="new_subject" size="95%" value="<%=rs("THESIS_SUBJECT")%>" readonly /></td></tr>
<tr><td>作者姓名：&emsp;&emsp;&emsp;<input type="text" class="txt full-width" name="author" size="18" value="<%=rs("STU_NAME")%>" readonly />&nbsp;
学号：<input type="text" class="txt full-width" name="stuno" size="15" value="<%=rs("STU_NO")%>" readonly />&nbsp;
学位类别：<input type="text" class="txt full-width" name="degreename" size="10" value="<%=rs("TEACHTYPE_NAME")%>" readonly /></td></tr>
<tr><td>指导教师：&emsp;&emsp;&emsp;<input type="text" class="txt full-width" name="tutorname" size="95%" value="<%=rs("TUTOR_NAME")%>" readonly /></td></tr><%
	If reviewfile_type=2 Then %>
<tr><td>领域名称：&emsp;&emsp;&emsp;<input type="text" class="txt full-width" name="speciality" size="95%" value="<%=rs("SPECIALITY_NAME")%>" readonly /></td></tr><%
	End If %>
<tr><td>研究方向：&emsp;&emsp;&emsp;<input type="text" class="txt full-width" name="researchway_name" size="95%" value="<%=rs("RESEARCHWAY_NAME")%>" readonly /></td></tr>
<tr><td>院系名称：&emsp;&emsp;&emsp;<input type="text" class="txt full-width" name="faculty" size="30%" value="工商管理学院" readonly />&nbsp;
班级：<input type="text" class="txt full-width" name="class" size="51%" value="<%=rs("CLASS_NAME")%>" readonly /></td></tr><%
	If Not IsNull(rs("THESIS_FORM")) And Len(rs("THESIS_FORM")) Then %>
<tr><td>论文形式：&emsp;&emsp;&emsp;<input type="text" class="txt full-width" name="thesisform" size="95%" value="<%=rs("THESIS_FORM")%>" readonly /></td></tr><%
	End If %>
<tr><td>开题报告表：&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile1" size="100" /></td></tr>
<tr><td>开题论文：&emsp;&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile2" size="100" /></td></tr>
<tr><td>中期检查表：&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile3" size="100" /></td></tr>
<tr><td>中期论文：&emsp;&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile4" size="100" /></td></tr>
<tr><td>预答辩意见书：&emsp;&emsp;&emsp;<input type="file" name="uploadfile5" size="100" /></td></tr>
<tr><td>预答辩论文：&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile6" size="100" /></td></tr>
<tr><td>送检论文：&emsp;&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile8" size="100" /></td></tr>
<tr><td>送检论文检测报告：&emsp;<input type="file" name="uploadfile13" size="100" /></td></tr>
<tr><td>送审论文：&emsp;&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile9" size="100" /></td></tr>
<tr><td>答辩论文：&emsp;&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile10" size="100" /></td></tr>
<tr><td>教指委盲评论文：&emsp;&emsp;<input type="file" name="uploadfile11" size="100" /></td></tr>
<tr><td>定稿论文：&emsp;&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile12" size="100" /></td></tr>
<tr><td>答辩审批材料：&emsp;&emsp;&emsp;<input type="file" name="uploadfile7" size="100" /></td></tr>
<tr><td>更改表格审核状态：&emsp;<select name="new_task_progress"><%
GetMenuListPubTerm "ReviewStatuses","STATUS_ID1","STATUS_NAME",task_progress,"AND STATUS_ID1 IS NOT NULL"
%></select></td></tr>
<tr><td>更改论文审核状态：&emsp;<select name="new_review_status"><%
GetMenuListPubTerm "ReviewStatuses","STATUS_ID2","STATUS_NAME",review_status,"AND STATUS_ID2 IS NOT NULL"
%></select></td></tr>
<tr class="buttons">
<td><p align="center"><input type="button" id="btnsubmit" name="btnsubmit" value="提 交" />&emsp;
<input type="button" value="返 回" onclick="history.go(-1)" />&emsp;
<input type="button" value="返回论文列表" onclick="document.all.ret.submit()" />
</p></td></tr></table>
<input type="hidden" name="In_ActivityId2" value="<%=activity_id%>">
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID2" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form></center>
<form id="ret" name="ret" action="paperList.asp" method="post">
<input type="hidden" name="In_ActivityId" value="<%=activity_id%>">
<input type="hidden" name="In_TEACHTYPE_ID" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize" value="<%=pageSize%>" />
<input type="hidden" name="pageNo" value="<%=pageNo%>" /></form>
</body><script type="text/javascript">
	document.all.btnsubmit.onclick=function() {
		this.value="正在提交，请稍候……";
		this.disabled=true;
		this.form.submit();
	}
	document.all.btnsubmit.disabled=false;
</script></html><%
Case 2	' 文件上传页面

	Dim Upload
	Set Upload=New ExtendedRequest
	
	new_task_progress=Upload.Form("new_task_progress")
	new_review_status=Upload.Form("new_review_status")
	activity_id=Upload.Form("In_ActivityId2")
	teachtype_id=Upload.Form("In_TEACHTYPE_ID2")
	class_id=Upload.Form("In_CLASS_ID2")
	enter_year=Upload.Form("In_ENTER_YEAR2")
	query_task_progress=Upload.Form("In_TASK_PROGRESS2")
	query_review_status=Upload.Form("In_REVIEW_STATUS2")
	finalFilter=Upload.Form("finalFilter2")
	pageSize=Upload.Form("pageSize2")
	pageNo=Upload.Form("pageNo2")
	
	Dim conn,rs,sql,sqlDetect,count
	sqlDetect=""
	Connect conn
	sql="SELECT * FROM Dissertations WHERE ID="&paper_id
	GetRecordSet conn,rs,sql,count
	If rs.EOF Then
	%><body><center><font color=red size="4">数据库没有该论文记录！</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	  CloseRs rs
	  CloseConn conn
		Response.End()
	End If
	
	Dim upFile,fso
	Dim msg
	Randomize()
	
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	For i=1 To UBound(arrDefaultFileListName)
		Set upFile=Upload.File("uploadfile"&i)
		If Len(upFile.FileName) Then
			' 检查上传目录是否存在
			uploadPath=Server.MapPath(arrDefaultFileListPath(i))
			If Not fso.FolderExists(uploadPath) Then fso.CreateFolder(uploadPath)
			file_ext=LCase(upFile.FileExt)
			' 生成日期格式文件名
			fileid=FormatDateTime(Now(),1)&Int(Timer)
			destFile=fileid&"."&file_ext
			destPath=uploadPath&"\"&destFile
			' 保存
			upFile.SaveAs destPath
			Select Case i
			Case 8	' 送检论文
				If rs("REVIEW_STATUS")=rsDetectPaperUploaded Then
					sqlDetect=sqlDetect&"EXEC spDeleteDetectResult "&rs("ID")&","&toSqlString(rs("THESIS_FILE"))&";"
				End If
				sqlDetect=sqlDetect&"EXEC spAddDetectResult "&paper_id&","&toSqlString(destFile)&",NULL,NULL,NULL,1;"
				rs("THESIS_FILE")=destFile
			Case 11	' 教指委盲评论文
				If rs("REVIEW_STATUS")=rsInstructReviewPaperUploaded Then
					sqlDetect=sqlDetect&"EXEC spDeleteDetectResult "&rs("ID")&","&toSqlString(rs("THESIS_FILE4"))&";"
				End If
				sqlDetect=sqlDetect&"EXEC spAddDetectResult "&paper_id&","&toSqlString(destFile)&",NULL,NULL,NULL,2;"
				rs("THESIS_FILE4")=destFile
			Case 13	' 送检论文检测报告
				sqlDetect=sqlDetect&"EXEC spSetDetectResultReport "&paper_id&","&toSqlString(rs("THESIS_FILE"))&","&toSqlString(destFile)&";"
			Case Else
				rs(arrDefaultFileListField(i))=destFile
			End Select
			
			msg=msg&arrDefaultFileListName(i)&"已上传成功并链接到论文的数据库记录。"&vbNewLine
		End If
	Next
	Set fso=Nothing
	If Len(new_task_progress) Then
		rs("TASK_PROGRESS")=new_task_progress
	End If
	If Len(new_review_status) Then
		rs("REVIEW_STATUS")=new_review_status
	End If
	rs.Update()
	If Len(sqlDetect) Then
		conn.Execute sqlDetect
	End If
%><form id="ret" action="paperDetail.asp?tid=<%=paper_id%>" method="post">
<input type="hidden" name="In_ActivityId2" value="<%=activity_id%>">
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID2" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form>
<script type="text/javascript">
	alert("操作完成，操作结果如下：\r\n<%=toJsString(msg)%>");
	document.all.ret.submit();
</script><%
End Select
CloseRs rs
CloseConn conn
%>
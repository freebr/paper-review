<%Response.Charset="utf-8"%>
<!--#include file="../inc/upload_5xsoft.inc"-->
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("user")) Then Response.Redirect("../error.asp?timeout")
thesisID=Request.QueryString("tid")
curstep=Request.QueryString("step")

Select Case curstep
Case vbNullString	' 论文详情页面
	period_id=Request.Form("In_PERIOD_ID2")
	teachtype_id=Request.Form("In_TEACHTYPE_ID2")
	class_id=Request.Form("In_CLASS_ID2")
	enter_year=Request.Form("In_ENTER_YEAR2")
	query_task_progress=Request.Form("In_TASK_PROGRESS2")
	query_review_status=Request.Form("In_REVIEW_STATUS2")
	finalFilter=Request.Form("finalFilter2")
	pageSize=Request.Form("pageSize2")
	pageNo=Request.Form("pageNo2")
	If Len(thesisID)=0 Or Not IsNumeric(thesisID) Then
	%><body bgcolor="ghostwhite"><center><font color=red size="4">参数无效。</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		Response.End
	End If
	Dim table_file(4)
	Connect conn
	sql="SELECT *,dbo.getThesisStatusText(1,TASK_PROGRESS,1) AS STAT_TEXT1,dbo.getThesisStatusText(2,REVIEW_STATUS,1) AS STAT_TEXT2 FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE ID="&thesisID
	GetRecordSet conn,rs,sql,result
	If result=0 Then
	%><body bgcolor="ghostwhite"><center><font color=red size="4">数据库没有该论文记录！</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	  CloseRs rs
	  CloseConn conn
		Response.End
	End If
	
	tutor_id=rs("TUTOR_ID")
	task_progress=rs("TASK_PROGRESS")
	review_status=rs("REVIEW_STATUS")
	For i=1 To 4
		table_file(i)=rs("TABLE_FILE"&i)
	Next
	If Not IsNull(rs("THESIS_FILE")) Then
		thesis_file=rs("THESIS_FILE")
	End If
	If Not IsNull(rs("THESIS_FILE2")) Then
		thesis_file_review=rs("THESIS_FILE2")
	End If
	If Not IsNull(rs("THESIS_FILE3")) Then
		thesis_file_modified=rs("THESIS_FILE3")
	End If
	If Not IsNull(rs("THESIS_FILE4")) Then
		thesis_file_modified=rs("THESIS_FILE4")
	End If
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
<script src="../scripts/utils.js" type="text/javascript"></script>
<script src="../scripts/thesis.js" type="text/javascript"></script>
<meta name="theme-color" content="#2D79B2" />
<title>上传表格/论文文件</title>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>上传表格/论文文件</font>
<form id="fmDetail" action="?step=2&tid=<%=thesisID%>" enctype="multipart/form-data" method="post">
<table class="tblform" width="800" cellspacing="1" cellpadding="3">
<tr><td>论文题目：&emsp;&emsp;&emsp;<input type="text" class="txt" name="new_subject" size="95%" value="<%=rs("THESIS_SUBJECT")%>" readonly /></td></tr>
<tr><td>作者姓名：&emsp;&emsp;&emsp;<input type="text" class="txt" name="author" size="18" value="<%=rs("STU_NAME")%>" readonly />&nbsp;
学号：<input type="text" class="txt" name="stuno" size="15" value="<%=rs("STU_NO")%>" readonly />&nbsp;
学位类别：<input type="text" class="txt" name="degreename" size="10" value="<%=rs("TEACHTYPE_NAME")%>" readonly />&nbsp;
学期：<input type="text" class="txt" name="new_period_id" size="6" value="<%=rs("PERIOD_ID")%>" readonly /></td></tr>
<tr><td>指导教师：&emsp;&emsp;&emsp;<input type="text" class="txt" name="tutorname" size="95%" value="<%=rs("TUTOR_NAME")%>" readonly /></td></tr><%
	If reviewfile_type=2 Then %>
<tr><td>领域名称：&emsp;&emsp;&emsp;<input type="text" class="txt" name="speciality" size="95%" value="<%=rs("SPECIALITY_NAME")%>" readonly /></td></tr><%
	End If %>
<tr><td>研究方向：&emsp;&emsp;&emsp;<input type="text" class="txt" name="researchway_name" size="95%" value="<%=rs("RESEARCHWAY_NAME")%>" readonly /></td></tr>
<tr><td>院系名称：&emsp;&emsp;&emsp;<input type="text" class="txt" name="faculty" size="30%" value="工商管理学院" readonly />&nbsp;
班级：<input type="text" class="txt" name="class" size="51%" value="<%=rs("CLASS_NAME")%>" readonly /></td></tr><%
	If Not IsNull(rs("THESIS_FORM")) And Len(rs("THESIS_FORM")) Then %>
<tr><td>论文形式：&emsp;&emsp;&emsp;<input type="text" class="txt" name="thesisform" size="95%" value="<%=rs("THESIS_FORM")%>" readonly /></td></tr><%
	End If %>
<tr><td>开题报告表：&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile1" size="100" /></td></tr>
<tr><td>开题论文：&emsp;&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile2" size="100" /></td></tr>
<tr><td>中期检查表：&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile3" size="100" /></td></tr>
<tr><td>中期论文：&emsp;&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile4" size="100" /></td></tr>
<tr><td>预答辩申请表：&emsp;&emsp;&emsp;<input type="file" name="uploadfile5" size="100" /></td></tr>
<tr><td>预答辩论文：&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile6" size="100" /></td></tr>
<tr><td>送检论文：&emsp;&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile7" size="100" /></td></tr>
<tr><td>送检论文检测报告：&emsp;<input type="file" name="uploadfile8" size="100" /></td></tr>
<tr><td>送审论文：&emsp;&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile9" size="100" /></td></tr>
<tr><td>修改后论文：&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile10" size="100" /></td></tr>
<tr><td>定稿论文：&emsp;&emsp;&emsp;&emsp;&emsp;<input type="file" name="uploadfile11" size="100" /></td></tr>
<tr><td>答辩审批材料：&emsp;&emsp;&emsp;<input type="file" name="uploadfile12" size="100" /></td></tr>
<tr><td>更改表格审核状态：&emsp;<select name="new_task_progress"><%
GetMenuListPubTerm "CODE_THESIS_REVIEW_STATUS","STATUS_ID1","STATUS_NAME",task_progress,"AND STATUS_ID1 IS NOT NULL"
%></select></td></tr>
<tr><td>更改论文审核状态：&emsp;<select name="new_review_status"><%
GetMenuListPubTerm "CODE_THESIS_REVIEW_STATUS","STATUS_ID2","STATUS_NAME",review_status,"AND STATUS_ID2 IS NOT NULL"
%></select></td></tr>
<tr class="trbuttons">
<td><p align="center"><input type="button" id="btnsubmit" name="btnsubmit" value="提 交" />&emsp;
<input type="button" value="返 回" onclick="history.go(-1)" />&emsp;
<input type="button" value="返回论文列表" onclick="document.all.ret.submit()" />
</p></td></tr></table>
<input type="hidden" name="In_PERIOD_ID2" value="<%=period_id%>">
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID2" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form></center>
<form id="ret" name="ret" action="thesisList.asp" method="post">
<input type="hidden" name="In_PERIOD_ID" value="<%=period_id%>">
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
	Set Upload=New upload_5xsoft
	
	new_task_progress=Upload.Form("new_task_progress")
	new_review_status=Upload.Form("new_review_status")
	period_id=Upload.Form("In_PERIOD_ID2")
	teachtype_id=Upload.Form("In_TEACHTYPE_ID2")
	class_id=Upload.Form("In_CLASS_ID2")
	enter_year=Upload.Form("In_ENTER_YEAR2")
	query_task_progress=Upload.Form("In_TASK_PROGRESS2")
	query_review_status=Upload.Form("In_REVIEW_STATUS2")
	finalFilter=Upload.Form("finalFilter2")
	pageSize=Upload.Form("pageSize2")
	pageNo=Upload.Form("pageNo2")
	
	Dim conn,rs,sql,result
	Connect conn
	sql="SELECT * FROM TEST_THESIS_REVIEW_INFO WHERE ID="&thesisID
	GetRecordSet conn,rs,sql,result
	If rs.EOF Then
	%><body bgcolor="ghostwhite"><center><font color=red size="4">数据库没有该论文记录！</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	  CloseRs rs
	  CloseConn conn
		Response.End
	End If
	
	Dim arrFileListName:arrFileListName=Array("","开题报告表","开题论文","中期检查表","中期论文","预答辩申请表","预答辩论文","答辩及授予学位审批材料","送检论文","送审论文","修改后论文","定稿论文","送检论文检测报告")
	Dim arrFileListPath:arrFileListPath=Array("","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/admin/upload/report")
	Dim arrFileListField:arrFileListField=Array("","TABLE_FILE1","TBL_THESIS_FILE1","TABLE_FILE2","TBL_THESIS_FILE2","TABLE_FILE3","TBL_THESIS_FILE3","TABLE_FILE4","THESIS_FILE","THESIS_FILE2","THESIS_FILE3","THESIS_FILE4","DETECT_REPORT")
	Dim upFile,fso
	Dim msg
	Randomize
	
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	For i=1 To UBound(arrFileListName)
		Set upFile=Upload.File("uploadfile"&i)
		If upFile.FileName<>vbNullString Then
			' 检查上传目录是否存在
			strUploadPath=Server.MapPath(arrFileListPath(i))
			If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)
			fileExt=LCase(upFile.FileExt)
			' 生成日期格式文件名
			fileid=FormatDateTime(Now(),1)&Int(Timer)
			strDestFile=fileid&"."&fileExt
			strDestPath=strUploadPath&"\"&strDestFile
			' 保存
			upFile.SaveAs strDestPath
			rs(arrFileListField(i))=strDestFile
			
			msg=msg&arrFileListName(i)&"已上传成功并链接到论文的数据库记录。"&vbNewLine
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
%><form id="ret" action="thesisDetail.asp?tid=<%=thesisID%>" method="post">
<input type="hidden" name="In_PERIOD_ID2" value="<%=period_id%>">
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
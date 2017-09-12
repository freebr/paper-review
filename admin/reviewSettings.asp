<%Response.Charset="utf-8"%>
<!--#include file="../inc/upload_5xsoft.inc"-->
<!--#include file="../inc/db.asp"-->
<%If IsEmpty(Session("user")) Then Response.Redirect("../error.asp?timeout")
Dim conn,rs,sql,result
Connect conn
curStep=Request.QueryString("step")
Select Case curStep
Case vbNullstring ' 设置页面
	sql="SELECT * FROM CODE_REVIEW_TYPE"
	GetRecordSetNoLock conn,rs,sql,result
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>评阅书类型设置</b></font>
<form id="fmReview" action="?step=1" method="POST" enctype="multipart/form-data">
<table width="900" cellpadding="2" cellspacing="1" bgcolor="dimgray"><tbody id="tbItems">
<tr bgcolor="ghostwhite"><td>当前共有&nbsp;<span id="spanNumItems" style="font-weight:bold">0</span>&nbsp;个条目</td></tr>
<tr id="trPanel" bgcolor="ghostwhite"><td>
<p><input type="button" name="btnadd" value="＋ 增加条目" onclick="addReviewTypeItem()" />&nbsp;
<input type="submit" name="btnsubmit" value="提交设置" /><input type="hidden" name="num_items" /><input type="hidden" name="num_olditems" value="<%=rs.RecordCount%>" />
</p></td></tr></tbody></table>
</form></center>
<script src="../scripts/reviewSettings.js" type="text/javascript"></script>
<script type="text/javascript"><%
	Do While Not rs.EOF %>
	addReviewTypeItem(<%=rs("ID")%>,'<%=toJsString(rs("TYPE_NAME"))%>',<%=toJsString(rs("TEACHTYPE_ID"))%>,'<%=toJsString(rs("THESIS_FORM"))%>','<%=toJsString(rs("REVIEW_FILE"))%>');<%
		rs.MoveNext()
	Loop %>
</script></body></html><%
	CloseRs rs
Case 1	' 后台进程

	Dim numItems,numOldItems
	Dim type_name,teachtype_id,thesis_form,file_id
	Dim rids,arr_rid,arr_typename,arr_teachtypeid,arr_thesisform,arr_fileid
	Dim fso,Upload,file,bError,delim
	
	Set Upload=New upload_5xsoft
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	' 检查上传目录是否存在
	strUploadPath = Server.MapPath("upload/review")
	If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)
	
	numItems=Int(Upload.Form("num_items"))
	numOldItems=Int(Upload.Form("num_olditems"))
	
	byteFileSize=0
	ReDim arr_fileid(numItems)
	Randomize
	For i=1 To numItems
		Set file=Upload.File("reviewFile"&i)
		If Len(file.FileName)=0 Then
			If i>numOldItems Then
				bError=True
				errdesc="请为第&nbsp;"&i&"&nbsp;个条目上传评阅书模板文件！"
				Exit For
			End If
		Else
			fileExt=LCase(file.FileExt)
			'If fileExt<>"pdf" Then
				'bError=True
				'errdesc="请为第&nbsp;"&i&"&nbsp;个条目上传PDF格式文件！"
				'Exit For
			'End If
			' 生成日期格式文件名
			fileid = FormatDateTime(Now(),1)&Int(Timer)&Int(Rnd()*999)
			strDestFile = fileid&"."&fileExt
			strDestPath = strUploadPath&"\"&strDestFile
			byteFileSize = byteFileSize+file.FileSize
			' 保存
			file.SaveAs strDestPath
			arr_fileid(i-1)=strDestFile
		End If
		Set file=Nothing
	Next
	Set fso=Nothing
	
	If bError Then
%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		Response.End
	End If
	
	' 获取表单上旧条目的ID
	delim=", "
	rids=Upload.Form("rid")
	arr_rid=Split(rids,delim)
	arr_typename=Split(Upload.Form("typename"),delim)
	arr_teachtypeid=Split(Upload.Form("teachtypeid"),delim)
	arr_thesisform=Split(Upload.Form("thesisform"),delim)
	If Len(rids)=0 Then rids="0"
	' 删除不显示在表单上的旧条目
	sql="DELETE FROM CODE_REVIEW_TYPE WHERE ID NOT IN ("&rids&")"
	conn.Execute sql
	
	sql="SELECT * FROM CODE_REVIEW_TYPE"
	GetRecordSet conn,rs,sql,result
	For i=0 To numItems-1
		type_name=arr_typename(i)
		teachtype_id=arr_teachtypeid(i)
		thesis_form=arr_thesisform(i)
		file_id=arr_fileid(i)
		If i<numOldItems Then
			' 更新记录
			rs.Find("ID="&arr_rid(i))
		Else
			' 添加记录
			rs.AddNew()
		End If
		rs("TYPE_NAME")=type_name
		rs("TEACHTYPE_ID")=teachtype_id
		rs("THESIS_FORM")=thesis_form
		If Len(file_id) Then	' 上传新文件
			rs("REVIEW_FILE")=file_id
		End If
	Next
	rs.UpdateBatch()
	CloseRs rs
	Set Upload=Nothing
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>评阅书类型设置</title>
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite">
<center><br /><b>评阅书类型设置</b><br /><br />
<form id="fmReview" action="reviewSettings.asp" method="POST">
<p><%=byteFileSize%> 字节已上传。</p></form>
<script type="text/javascript">
	alert("操作完成。");
	$('#fmReview').submit();
</script></center></body></html><%
End Select
CloseConn conn
%>
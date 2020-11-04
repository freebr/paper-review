<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
Dim conn,rs,sql,count
Connect conn
step=Request.QueryString("step")
Select Case step
Case vbNullstring ' 设置页面
	sql="SELECT * FROM ReviewTypes"
	GetRecordSetNoLock conn,rs,sql,count
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>评阅书类型设置</title>
<% useStylesheet "admin" %>
<% useScript "jquery", "common", "reviewSettings" %>
</head>
<body>
<center><font size=4><b>评阅书类型设置</b></font>
<form id="fmReview" action="?step=1" method="POST" enctype="multipart/form-data">
<table width="900" cellpadding="2" cellspacing="1" bgcolor="dimgray"><tbody id="tbItems">
<tr bgcolor="ghostwhite"><td>当前共有&nbsp;<span id="spanNumItems" style="font-weight:bold">0</span>&nbsp;个条目</td></tr>
<tr id="trPanel" bgcolor="ghostwhite"><td>
<p><input type="button" name="btnadd" value="＋ 增加条目" onclick="addReviewTypeItem()" />&nbsp;
<input type="submit" name="btnsubmit" value="提交设置" /><input type="hidden" name="num_items" /><input type="hidden" name="num_olditems" value="<%=rs.RecordCount%>" />
</p></td></tr></tbody></table>
</form></center>
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
	Dim Upload,file,bError,delim
	
	Set Upload=New ExtendedRequest
	ensurePathExists Server.MapPath(uploadBasePath(usertypeAdmin,"review_template"))
	
	numItems=Int(Upload.Form("num_items"))
	numOldItems=Int(Upload.Form("num_olditems"))
	
	byteFileSize=0
	ReDim arr_fileid(numItems)
	Randomize()
	For i=1 To numItems
		Set file=Upload.File("reviewFile"&i)
		If Len(file.FileName)=0 Then
			If i>numOldItems Then
				bError=True
				errMsg="请为第&nbsp;"&i&"&nbsp;个条目上传评阅书模板文件！"
				Exit For
			End If
		Else
			file_ext=LCase(file.FileExt)
			'If file_ext<>"pdf" Then
				'bError=True
				'errMsg="请为第&nbsp;"&i&"&nbsp;个条目上传PDF格式文件！"
				'Exit For
			'End If
			destFile = timestamp()&Int(Rnd()*999)&"."&file_ext
			destPath = strUploadPath&"\"&destFile
			byteFileSize = byteFileSize+file.FileSize
			' 保存
			file.SaveAs destPath
			arr_fileid(i-1)=destFile
		End If
		Set file=Nothing
	Next
	
	If bError Then
		showErrorPage errMsg, "提示"
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
	sql="DELETE FROM ReviewTypes WHERE ID NOT IN ("&rids&")"
	conn.Execute sql
	
	sql="SELECT * FROM ReviewTypes"
	GetRecordSet conn,rs,sql,count
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
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body>
<script type="text/javascript">
	alert("操作完成。");
	location.href="reviewSettings.asp";
</script></body></html><%
End Select
CloseConn conn
%>
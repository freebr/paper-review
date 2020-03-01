﻿<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("StuId")) Then Response.Redirect("../error.asp?timeout")
thesisID=Request.QueryString("tid")
filetype=Request.QueryString("type")
If Not IsNumeric(filetype) Then
	bError=True
	errdesc="参数无效。"
Else
	filetype=Int(filetype)
	If filetype<1 Or filetype>UBound(arrFileListName) Then
		bError=True
		errdesc="参数无效。"
	End If
End If
If bError Then
	showErrorPage errdesc, "提示"
End If

Connect conn
sql="SELECT * FROM ViewDissertations_student WHERE ID="&thesisID
GetRecordSetNoLock conn,rs,sql,count
If rs.EOF Then
	CloseRs rs
	CloseConn conn
	showErrorPage "数据库没有该论文记录！", "提示"
End If

Dim source_file,file_ext,newfilename
Dim bReviewFileVisible(1)
Dim fso,file,stream
Set fso=Server.CreateObject("Scripting.FileSystemObject")
source_file=rs(arrFileListField(filetype))
bReviewFileVisible(0)=(rs("ReviewFileDisplayStatus1") And 2)<>0
bReviewFileVisible(1)=(rs("ReviewFileDisplayStatus2") And 2)<>0
If IsNull(source_file) Then
	source_file=""
Else
	file_ext=LCase(fso.GetExtensionName(source_file))
	If filetype=17 Or filetype=18 Then ' 评阅书则提供无专家信息版本
		' 根据评阅书显示设置决定是否显示文件
		If Not bReviewFileVisible(filetype-17) Then
			source_file=""
		Else
			source_file=arrFileListPath(filetype)&"/"&fso.GetBaseName(source_file)&"_noexp."&file_ext
		End If
	Else
		source_file=arrFileListPath(filetype)&"/"&source_file
	End If
	source_file=Server.MapPath(baseUrl()&source_file)
End If
If Not fso.FileExists(source_file) Then
	Set fso=Nothing
	showErrorPage "该论文暂无"&arrFileListName(filetype)&"或已被删除！", "提示"
End If
Set file=fso.GetFile(source_file)
If Len(arrFileListNamePostfix(filetype)) Then
	newfilename=rs("SPECIALITY_NAME")&rs("STU_NAME")&rs("STU_NO")&"-"&arrFileListNamePostfix(filetype)
Else
	subject=Replace(rs("THESIS_SUBJECT"),":","_")
	subject=Replace(subject,"""","_")
	subject=Replace(subject,"<","_")
	subject=Replace(subject,">","_")
	subject=Replace(subject,"?","_")
	subject=Replace(subject,"\","_")
	subject=Replace(subject,"/","_")
	subject=Replace(subject,"|","_")
	subject=Replace(subject,"*","_")
	newfilename=rs("SPECIALITY_NAME")&"-"&subject
End If
newfilename=newfilename&"."&file_ext
Set stream=Server.CreateObject("ADODB.Stream")
stream.Mode=3
stream.Type=1
stream.Open()
stream.LoadFromFile source_file
Response.Buffer=True
Response.Clear()
Session.Codepage=936
Response.AddHeader "Content-Disposition","attachment; filename="&newfilename
Response.AddHeader "Content-Type","application/octet-stream"
Response.AddHeader "Content-Length",file.Size
block_size=10240
Do While Not stream.EOS
	Response.BinaryWrite stream.Read(block_size)
	Response.Flush()
Loop
Session.Codepage=65001
stream.Close()
Set stream=Nothing
Set file=Nothing
Set fso=Nothing
CloseRs rs
CloseConn conn
%>
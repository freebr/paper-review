<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Tid")) Then Response.Redirect("../error.asp?timeout")
thesisID=Request.QueryString("tid")
filetype=Request.QueryString("type")
hash=Request.QueryString("hash")
If Not IsNumeric(filetype) Then
	bError=True
	errdesc="参数无效。"
Else
	filetype=Int(filetype)
	If filetype<1 Or filetype>UBound(arrDefaultFileListName) Then
		bError=True
		errdesc="参数无效。"
	End If
End If
If bError Then
	showErrorPage errdesc, "提示"
End If

Connect conn
sql="SELECT *,LEFT(REVIEW_FILE,CHARINDEX(',',REVIEW_FILE)-1) AS REVIEW_FILE1,RIGHT(REVIEW_FILE,LEN(REVIEW_FILE)-CHARINDEX(',',REVIEW_FILE)) AS REVIEW_FILE2 FROM ViewDissertations WHERE ID="&thesisID&" AND Valid=1"
GetRecordSetNoLock conn,rs,sql,count
If rs.EOF Then
	CloseRs rs
	CloseConn conn
	showErrorPage "数据库没有该论文记录！", "提示"
End If

Dim source_file,fileExt,newfilename
Dim fso,file,stream
Set fso=Server.CreateObject("Scripting.FileSystemObject")

If (filetype=8 Or filetype=13) And Len(hash) Then
	sql="SELECT * FROM ViewDetectResult WHERE THESIS_ID="&thesisID&" AND HASH="&toSqlString(hash)
	GetRecordSet conn,rsDetect,sql,count
	If filetype=8 Then
		source_file=rsDetect("THESIS_FILE").Value
	Else
		source_file=rsDetect("DETECT_REPORT").Value
	End If
	CloseRs rsDetect
Else
	source_file=rs(arrDefaultFileListField(filetype)).Value
End If
If IsNull(source_file) Then
	source_file=""
Else
	fileExt=LCase(fso.GetExtensionName(source_file))
	If filetype=15 Or filetype=16 Then ' 评阅书则提供无专家信息版本
		' 根据评阅书显示设置决定是否显示文件
		If (rs("REVIEW_FILE_STATUS") And 1)=0 Then
			source_file=arrDefaultFileListPath(filetype)
		Else
			source_file=arrDefaultFileListPath(filetype)&"/"&fso.GetBaseName(source_file)&"_noexp."&fileExt
		End If
	Else
		source_file=arrDefaultFileListPath(filetype)&"/"&source_file
	End If
	source_file=Server.MapPath(source_file)
End If
If Not fso.FileExists(source_file) Then
	Set fso=Nothing
	showErrorPage "该论文暂无"&arrDefaultFileListName(filetype)&"或已被删除！", "提示"
End If
Set file=fso.GetFile(source_file)
If Len(arrDefaultFileListNamePostfix(filetype)) Then
	newfilename=rs("SPECIALITY_NAME")&rs("STU_NAME")&rs("STU_NO")&"-"&arrDefaultFileListNamePostfix(filetype)
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
newfilename=newfilename&"."&fileExt
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
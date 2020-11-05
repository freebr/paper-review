<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")
paper_id=Request.QueryString("tid")
store=Request.QueryString("store")
filetype=Request.QueryString("type")
id=Request.QueryString("id")
If Not IsNumeric(filetype) Then
	bError=True
	errMsg="参数无效。"
Else
	filetype=Int(filetype)
	If filetype<1 Or filetype>UBound(arrDefaultFileListName) Then
		bError=True
		errMsg="参数无效。"
	End If
End If
If bError Then
	showErrorPage errMsg, "提示"
End If

ConnectDb conn

Dim source_file,file_ext,newfilename
Dim fso,file,stream
Set fso=CreateFSO()

If Not IsEmpty(store) Then
	Select Case store
		Case "audit"
			sql="SELECT * FROM AuditRecords WHERE Id=?"
			Set ret=ExecQuery(conn,sql,CmdParam("Id",adVarWChar,100,id))
			Set rsAudit=ret("rs")
			If rsAudit.EOF Then
				CloseRs rsAudit
				showErrorPage "找不到所需的审核记录！", "提示"
			End If
			paper_id=rsAudit("DissertationId")
			source_file=rsAudit("AuditFile")
			CloseRs rsAudit
		Case "detect"
			sql="SELECT * FROM DetectResults WHERE Id=?"
			Set ret=ExecQuery(conn,sql,CmdParam("Id",adVarWChar,100,id))
			Set rsDetect=ret("rs")
			If rsDetect.EOF Then
				CloseRs rsDetect
				showErrorPage "找不到所需的检测记录！", "提示"
			End If
			paper_id=rsDetect("DissertationId")
			If filetype=8 Then
				source_file=rsDetect("DetectFile")
			Else
				source_file=rsDetect("ReportFile")
			End If
			CloseRs rsDetect
		Case "review"
			sql="SELECT * FROM ViewReviewRecords WHERE Id=?"
			Set ret=ExecQuery(conn,sql,CmdParam("Id",adVarWChar,100,id))
			Set rsReview=ret("rs")
			If rsReview.EOF Then
				CloseRs rsReview
				showErrorPage "找不到所需的评阅记录！", "提示"
			End If
			paper_id=rsReview("DissertationId")
			source_file=rsReview("ReviewFile")&"_noexp.pdf"
			CloseRs rsReview
	End Select
End If
sql="SELECT * FROM ViewDissertations WHERE Id=?"
Set ret=ExecQuery(conn,sql,CmdParam("Id",adInteger,4,paper_id))
Set rs=ret("rs")
If rs.EOF Then
	CloseRs rs
	CloseConn conn
	showErrorPage "数据库没有该论文记录！", "提示"
End If
If IsEmpty(store) Then
	source_file=rs(arrDefaultFileListField(filetype))
End If

If IsNull(source_file) Then
	source_file=""
Else
	If filetype=16 Or filetype=17 Then ' 评阅书则提供无专家信息版本
		' 根据评阅书显示设置决定是否显示文件
		Dim bReviewFileVisible
		bReviewFileVisible=Array(rs("ReviewFileDisplayStatus1") > 0, rs("ReviewFileDisplayStatus2") > 0)
		If Not bReviewFileVisible(filetype-16) Then
			source_file=""
		Else
			source_file=resolvePath(arrDefaultFileListPath(filetype),fso.GetBaseName(source_file)&"_noexp.pdf")
		End If
	Else
		source_file=resolvePath(arrDefaultFileListPath(filetype),source_file)
	End If
	source_file=Server.MapPath(resolvePath(basePath(),source_file))
	file_ext=LCase(fso.GetExtensionName(source_file))
End If

If Not fso.FileExists(source_file) Then
	Set fso=Nothing
	showErrorPage "该论文暂无"&arrDefaultFileListName(filetype)&"或已被删除！", "提示"
End If
Set file=fso.GetFile(source_file)

If Len(arrDefaultFileListNamePostfix(filetype)) Then
	newfilename=rs("STU_NAME")&"_"&rs("STU_NO")&"_"&arrDefaultFileListNamePostfix(filetype)
Else
	subject=toFilenameString(rs("THESIS_SUBJECT").Value)
	newfilename=rs("STU_NAME")&"_"&rs("STU_NO")&"_"&subject
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
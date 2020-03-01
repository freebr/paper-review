<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
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
sql="SELECT * FROM ViewDissertations_admin WHERE ID="&thesisID&" AND Valid=1"
GetRecordSet conn,rs,sql,count
If rs.EOF Then
	CloseRs rs
	CloseConn conn
	showErrorPage "数据库没有该论文记录！", "提示"
End If

Dim source_file,file_ext,newfilename
Dim fso,file,stream
Set fso=Server.CreateObject("Scripting.FileSystemObject")

If (filetype=8 Or filetype=13) And Len(hash) Then
	sql="SELECT * FROM ViewDetectResults WHERE THESIS_ID=? AND HASH=?"
	Set ret=ExecQuery(conn,sql,_
		CmdParam("THESIS_ID",adInteger,4,thesisID),CmdParam("HASH",adVarWChar,100,hash))
	Set rsDetect=ret("rs")
	If filetype=8 Then
		source_file=rsDetect("THESIS_FILE")
	Else
		source_file=rsDetect("DETECT_REPORT")
	End If
	CloseRs rsDetect
ElseIf filetype=16 Or filetype=17 Then
	source_file=rs(arrDefaultFileListField(filetype))&".pdf"
ElseIf filetype=18 Then
	review_order=toUnsignedInt(Request.QueryString("order"))
	If review_order=-1 Then review_order=0
	sql="SELECT * FROM ViewReviewRecords WHERE DissertationId=? AND ReviewOrder=?"
	Set ret=ExecQuery(conn,sql,_
		CmdParam("DissertationId",adInteger,4,thesisID),_
		CmdParam("ReviewOrder",adInteger,4,review_order))
	Set rsReview=ret("rs")
	If rsReview.EOF Then
		showErrorPage "找不到评阅书！", "提示"
	End If
	source_file=rsReview("ReviewFile")&".pdf"
	CloseRs rsReview
Else
	source_file=rs(arrDefaultFileListField(filetype))
End If
If IsNull(source_file) Then
	source_file=""
Else
	source_file=Server.MapPath(baseUrl()&arrDefaultFileListPath(filetype)&"/"&source_file)
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
	subject=Replace(rs("THESIS_SUBJECT"),":","_")
	subject=Replace(subject,"""","_")
	subject=Replace(subject,"<","_")
	subject=Replace(subject,">","_")
	subject=Replace(subject,"?","_")
	subject=Replace(subject,"\","_")
	subject=Replace(subject,"/","_")
	subject=Replace(subject,"|","_")
	subject=Replace(subject,"*","_")
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
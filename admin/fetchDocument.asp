<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
Dim arrFileListName,arrFileListNamePostfix,arrFileListPath,arrFileListField
arrFileListName=Array("","开题报告表","开题论文","中期检查表","中期论文","预答辩申请表","预答辩论文","答辩及授予学位审批材料","送检论文","送审论文","答辩论文","定稿论文","送检论文检测报告","硕士学位论文送审申请表","论文评阅书 1","论文评阅书 2","论文评阅书")
arrFileListNamePostfix=Array("","开题报告表","开题论文","中期检查表","中期论文","预答辩申请表","预答辩论文","答辩审批材料","","","","","检测报告","送审审核表","论文评阅书(1)","论文评阅书(2)","论文评阅书")
arrFileListPath=Array("","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/admin/upload/report","/ThesisReview/tutor/export","/ThesisReview/expert/export","/ThesisReview/expert/export","/ThesisReview/expert/export")
arrFileListField=Array("","TABLE_FILE1","TBL_THESIS_FILE1","TABLE_FILE2","TBL_THESIS_FILE2","TABLE_FILE3","TBL_THESIS_FILE3","TABLE_FILE4","THESIS_FILE","THESIS_FILE2","THESIS_FILE3","THESIS_FILE4","DETECT_REPORT","REVIEW_APP","REVIEW_FILE1","REVIEW_FILE2")
thesisID=Request.QueryString("tid")
filetype=Request.QueryString("type")
hash=Request.QueryString("hash")
If Not IsNumeric(filetype) Then
	bError=True
	errdesc="参数无效。"
ElseIf filetype<1 Or filetype>16 Then
	bError=True
	errdesc="参数无效。"
End If
If bError Then
	showErrorPage errdesc, "提示"
End If

Connect conn
sql="SELECT *,LEFT(REVIEW_FILE,CHARINDEX(',',REVIEW_FILE)-1) AS REVIEW_FILE1,RIGHT(REVIEW_FILE,LEN(REVIEW_FILE)-CHARINDEX(',',REVIEW_FILE)) AS REVIEW_FILE2 FROM ViewDissertations WHERE ID="&thesisID&" AND Valid=1"
GetRecordSet conn,rs,sql,count
If rs.EOF Then
	CloseRs rs
	CloseConn conn
	showErrorPage "数据库没有该论文记录！", "提示"
End If

Dim source_file,fileExt,newfilename
Dim fso,file,stream
Set fso=Server.CreateObject("Scripting.FileSystemObject")

If (filetype=8 Or filetype=12) And Len(hash) Then
	sql="SELECT * FROM ViewDetectResult WHERE THESIS_ID=? AND HASH=?"
	Set ret=ExecQuery(conn,sql,_
		CmdParam("THESIS_ID",adInteger,4,thesisID),CmdParam("HASH",adVarWChar,100,hash))
	Set rsDetect=ret("rs")
	If filetype=8 Then
		source_file=rsDetect("THESIS_FILE").Value
	Else
		source_file=rsDetect("DETECT_REPORT").Value
	End If
	CloseRs rsDetect
ElseIf filetype=16 Then
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
	source_file=rsReview("ReviewFile").Value&".pdf"
	CloseRs rsReview
Else
	source_file=rs(arrFileListField(filetype)).Value
End If
If IsNull(source_file) Then
	source_file=""
Else
	fileExt=LCase(fso.GetExtensionName(source_file))
	source_file=Server.MapPath(arrFileListPath(filetype)&"/"&source_file)
End If

If Not fso.FileExists(source_file) Then
	Set fso=Nothing
	showErrorPage "该论文暂无"&arrFileListName(filetype)&"或已被删除！", "提示"
End If
Set file=fso.GetFile(source_file)
If Len(arrFileListNamePostfix(filetype)) Then
	newfilename=rs("STU_NAME").Value&"_"&rs("STU_NO").Value&"_"&arrFileListNamePostfix(filetype)
Else
	subject=Replace(rs("THESIS_SUBJECT").Value,":","_")
	subject=Replace(subject,"""","_")
	subject=Replace(subject,"<","_")
	subject=Replace(subject,">","_")
	subject=Replace(subject,"?","_")
	subject=Replace(subject,"\","_")
	subject=Replace(subject,"/","_")
	subject=Replace(subject,"|","_")
	subject=Replace(subject,"*","_")
	newfilename=rs("STU_NAME").Value&"_"&rs("STU_NO").Value&"_"&subject
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
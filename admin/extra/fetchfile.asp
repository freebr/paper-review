<%Response.Charset="utf-8"%>
<!--#include file="../../inc/db.asp"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("user")) Then Response.Redirect("../../error.asp?timeout")
Dim arrFileListName,arrFileListNamePostfix,arrFileListPath,arrFileListField
arrFileListName=Array("","送审论文","论文评阅书 1","论文评阅书 2")
arrFileListNamePostfix=Array("","","论文评阅书(1)","论文评阅书(2)")
arrFileListPath=Array("","/ThesisReview/student/upload","/ThesisReview/expert/export","/ThesisReview/expert/export")
arrFileListField=Array("","THESIS_FILE2","REVIEW_FILE1","REVIEW_FILE2")
thesisID=Request.QueryString("tid")
filetype=Request.QueryString("type")
If Not IsNumeric(filetype) Then
	bError=True
	errdesc="参数无效。"
ElseIf filetype<1 Or filetype>3 Then
	bError=True
	errdesc="参数无效。"
End If
If bError Then
%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="关 闭" onclick="window.close()" /></center></body><%
	Response.End
End If

Connect conn
sql="SELECT *,LEFT(REVIEW_FILE,CHARINDEX(',',REVIEW_FILE)-1) AS REVIEW_FILE1,RIGHT(REVIEW_FILE,LEN(REVIEW_FILE)-CHARINDEX(',',REVIEW_FILE)) AS REVIEW_FILE2 FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE ID="&thesisID&" AND Valid=1"
GetRecordSetNoLock conn,rs,sql,result
If result<>1 Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">数据库没有该论文记录！</font><br /><input type="button" value="关 闭" onclick="window.close()" /></center></body><%
	Response.End
End If

Dim sourcefile,fileExt,newfilename
Dim fso,file,stream
Set fso=Server.CreateObject("Scripting.FileSystemObject")
sourcefile=rs(arrFileListField(filetype))
If IsNull(sourcefile) Then
	sourcefile=""
Else
	fileExt=LCase(fso.GetExtensionName(sourcefile))
	If filetype=2 Or filetype=3 Then ' 评阅书则提供无学生信息版本
		sourcefile=arrFileListPath(filetype)&"/"&fso.GetBaseName(sourcefile)&"_nostu."&fileExt
	Else
		sourcefile=arrFileListPath(filetype)&"/"&sourcefile
	End If
	sourcefile=Server.MapPath(sourcefile)
End If
If Not fso.FileExists(sourcefile) Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">该论文暂无<%=arrFileListName(filetype)%>或已被删除！</font><br /><input type="button" value="关 闭" onclick="window.close()" /></center></body><%
	Set fso=Nothing
	Response.End
End If
Set file=fso.GetFile(sourcefile)
If Len(arrFileListNamePostfix(filetype)) Then
	newfilename=rs("SPECIALITY_NAME")&"-"&arrFileListNamePostfix(filetype)
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
stream.LoadFromFile sourcefile
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
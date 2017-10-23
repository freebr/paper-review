<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"-->
<%'If IsEmpty(Session("user")) Then Response.Redirect("../error.asp?timeout")
Dim arrFileListName,arrFileListNamePostfix,arrFileListPath,arrFileListField
arrFileListName=Array("","开题报告表","开题论文","中期检查表","中期论文","预答辩申请表","预答辩论文","答辩及授予学位审批材料","送检论文","送审论文","答辩论文","定稿论文","送检论文检测报告","硕士学位论文送审申请表","论文评阅书 1","论文评阅书 2")
arrFileListNamePostfix=Array("","开题报告表","开题论文","中期检查表","中期论文","预答辩申请表","预答辩论文","答辩审批材料","","","","","检测报告","送审审核表","论文评阅书(1)","论文评阅书(2)")
arrFileListPath=Array("","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/student/upload","/ThesisReview/admin/upload/report","/ThesisReview/teacher/export","/ThesisReview/expert/export","/ThesisReview/expert/export")
arrFileListField=Array("","TABLE_FILE1","TBL_THESIS_FILE1","TABLE_FILE2","TBL_THESIS_FILE2","TABLE_FILE3","TBL_THESIS_FILE3","TABLE_FILE4","THESIS_FILE","THESIS_FILE2","THESIS_FILE3","THESIS_FILE4","DETECT_REPORT","REVIEW_APP","REVIEW_FILE1","REVIEW_FILE2")
thesisID=Request.QueryString("tid")
filetype=Request.QueryString("type")
If Not IsNumeric(filetype) Then
	bError=True
	errdesc="参数无效。"
ElseIf filetype<1 Or filetype>15 Then
	bError=True
	errdesc="参数无效。"
End If
If bError Then
%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="关 闭" onclick="window.close()" /></center></body><%
	Response.End
End If

Connect conn
sql="SELECT *,LEFT(REVIEW_FILE,CHARINDEX(',',REVIEW_FILE)-1) AS REVIEW_FILE1,RIGHT(REVIEW_FILE,LEN(REVIEW_FILE)-CHARINDEX(',',REVIEW_FILE)) AS REVIEW_FILE2 FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE ID="&thesisID&" AND Valid=1"
GetRecordSet conn,rs,sql,result
If rs.EOF Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">数据库没有该论文记录！</font><br /><input type="button" value="关 闭" onclick="window.close()" /></center></body><%
	Response.End
End If

Dim sourcefile,fileExt,newfilename
Dim fso,file,stream
Set fso=Server.CreateObject("Scripting.FileSystemObject")
sourcefile=rs(arrFileListField(filetype))
'Response.Redirect arrFileListPath(filetype)&"/"&sourcefile
If IsNull(sourcefile) Then
	sourcefile=""
Else
	fileExt=LCase(fso.GetExtensionName(sourcefile))
	sourcefile=Server.MapPath(arrFileListPath(filetype)&"/"&sourcefile)
End If
If Not fso.FileExists(sourcefile) Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">该论文暂无<%=arrFileListName(filetype)%>或已被删除！</font><br /><input type="button" value="关 闭" onclick="window.close()" /></center></body><%
	Set fso=Nothing
	Response.End
End If
Set file=fso.GetFile(sourcefile)
If Len(arrFileListNamePostfix(filetype)) Then
	newfilename=rs("STU_NAME")&"_"&rs("STU_NO")&"_"&arrFileListNamePostfix(filetype)
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
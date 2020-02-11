<%Response.Buffer=True
Server.ScriptTimeout=10000%>
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
Dim uppath,fso,folder,files
Dim arr:arr=Split("THESIS_FILE,THESIS_FILE2,THESIS_FILE3,THESIS_FILE4,TABLE_FILE1,TABLE_FILE2,TABLE_FILE3,TABLE_FILE4,TBL_THESIS_FILE1,TBL_THESIS_FILE2,TBL_THESIS_FILE3",",")
Dim conn,rs,sql,count

uppath=Server.MapPath("/ThesisReview/student/upload")
startdate=Request.QueryString("from")
If IsEmpty(startdate) Then startdate="1900-1-1"
startdate=CDate(startdate)
enddate=Request.QueryString("to")
If IsEmpty(enddate) Then
	enddate=Date()
	enddatestr="至今"
Else
	enddate=CDate(enddate)
	enddatestr="至 "&enddate&" "
End If
enddate=DateAdd("d",enddate,1)
Set fso=Server.CreateObject("Scripting.FileSystemObject")
Set folder=fso.GetFolder(uppath)
Set files=folder.Files

Connect conn
%>
<html>
<head>
	<title>清查废旧上传文件</title>
</head>
<body bgcolor="black"><p style="color:yellow"><%
Response.Write "正在清查自 "&startdate&" "&enddatestr&"上传的废旧文件……<br/>"
del_count=0
field_str=Join(arr,",")
For Each file In files
	If DateDiff("d",startdate,file.DateCreated)>=0 And DateDiff("d",file.DateCreated,enddate)>0 Then
		sql="SELECT STU_NAME,STU_ID FROM ViewDissertations WHERE '"&file.Name&"' IN ("&field_str&")"
		GetRecordSetNoLock conn,rs,sql,count
		If Not rs.EOF Then
			Response.Write "文件 "&file.Name&" 正由学生["&rs(0)&"]("&rs(1)&")使用。<br/>"
		Else
			Response.Write "文件 "&file.Name&" 无用户使用，将删除。<br/>"
			fso.DeleteFile uppath&"\"&file.Name
			del_count=del_count+1
		End If
		Response.Flush()
		CloseRs rs
	End If
Next
CloseConn conn
Response.Write "清查完毕，共删除 "&del_count&" 个废旧文件。"
%></p></body></html>
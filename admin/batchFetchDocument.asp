<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
filetype=Request.Form("filetype")
ids=Request("sel")
numRecord=UBound(Split(ids,","))+1
step=Request.QueryString("step")
Select Case step
Case vbNullString	' 选择页面
	rarFilenamePostfix="(共"&numRecord&"份)"
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>批量下载表格/论文</title>
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>批量下载表格/论文</b><br />
<form action="?step=2" method="POST">
<p>您选择了&nbsp;<%=numRecord%>&nbsp;条论文记录</p>
<p>请选择要下载的文件：<select name="filetype"><option value="0">请选择</option><%
For i=1 To UBound(arrDefaultFileListName) %>
<option value="<%=i%>"><%=arrDefaultFileListName(i)%></option><%
Next
%></select></p>
<p>打包压缩文件名：<input type="text" name="rarfilename" size="40" />.rar&nbsp;</p>
<input type="hidden" name="sel" value="<%=ids%>" />
<input type="submit" name="btnsubmit" value="批量下载" />&nbsp;
<input type="button" name="btnclose" value="关 闭" onclick="tabmgr.removeTab(window.index)" /></p></form>
<p align="left"><span id="output" style="color:#000099;font-size:9pt"></span></p></center></body>
<script type="text/javascript">
$(document).ready(function(){
	var progfile="http://www.cnsba.com/PaperReview/admin/rar/tmp/prog_<%=Session("id")%>.txt";
	$('select').change(function() {
		if(!this.selectedIndex)return;
		$(':text').val(this.options[this.selectedIndex].innerText+"<%=rarFilenamePostfix%>");
	});
	$('form').submit(function() {
		$(':submit').val("正在处理，请稍候...")
			.attr('disabled',true);
		$('#output').html('');
		setTimeout(refreshProgress,500);
	});
	$(':submit').attr('disabled',false);
	function refreshProgress() {
		$.get(progfile,(data,status)=>{
			if(status=='success') {
				$('#output').html(data);
				if(/<ok\/>/.test(data)) {
					$(':submit').val('批量下载').attr('disabled',false);
				} else {
					setTimeout(refreshProgress,500);
				}
			} else {
				setTimeout(refreshProgress,500);
			}
		});
	}
});
</script></html><%
Case 2	' 下载页面

	numRecord=UBound(Split(ids,","))+1
	If numRecord=0 Then
		bError=True
		errdesc="请选择论文！"
	ElseIf Not IsNumeric(filetype) Then
		bError=True
		errdesc="参数无效。"
	Else
		filetype=Int(filetype)
		If filetype<1 Or filetype>UBound(arrDefaultFileListName) Then
			bError=True
			errdesc="请选择要批量下载的文件类型！"
		End If
	End If
	If bError Then
	%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		Response.End()
	End If
	
	rarFilename=Trim(Request.Form("rarfilename"))
	If LCase(Right(rarFilename,4))=".rar" Then rarFilename=Trim(Left(rarFilename,Len(rarFilename)-4))
	If Len(rarFilename)=0 Then	' 取默认文件名
		rarFilename=arrDefaultFileListName(filetype)&"(共"&numRecord&"份)"
	End If
	rarFilename=rarFilename&".rar"
	Connect conn
	sql="SELECT *,LEFT(REVIEW_FILE,CHARINDEX(',',REVIEW_FILE)-1) AS REVIEW_FILE1,RIGHT(REVIEW_FILE,LEN(REVIEW_FILE)-CHARINDEX(',',REVIEW_FILE)) AS REVIEW_FILE2 FROM ViewDissertations WHERE ID IN ("&ids&") AND Valid=1"
	GetRecordSet conn,rs,sql,count
	If rs.EOF Then
	%><body bgcolor="ghostwhite"><center><font color=red size="4">所选记录不存在！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		Response.End()
	End If
	' 打包文件
	Dim source_file,fileExt,oldfilename,newfilename
	Dim rarExe,rarFile,tmpDir,rarDir,sourcefilelist,renamefilelist,commentfile,progfile
	Dim comment,cmd
	Dim fso,streamLog,wsh
	Dim numFailed,numSucceeded,numCompleted
	
	numFailed=0
	numSucceeded=0
	numBatchSize=20
	rarExe=Server.MapPath("rar/Rar.exe")
	rarFile=Server.MapPath("rar/"&rarFilename)
	tmpDir=Server.MapPath("rar/tmp")
	rarDir=Server.MapPath("rar/tmp/"&rarFilename)
	commentfile=Server.MapPath("rar/tmp/comment_"&Timer&".txt")
	progfile=Server.MapPath("rar/tmp/prog_"&Session("id")&".txt")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	
	' 删除已有目录
	If fso.FolderExists(rarDir) Then fso.DeleteFolder rarDir
	fso.CreateFolder rarDir
	' 删除已有文件
	If fso.FileExists(rarFile) Then fso.DeleteFile rarFile
	If fso.FileExists(progfile) Then fso.DeleteFile progfile
	Set streamLog=Server.CreateObject("ADODB.Stream")
	streamLog.Mode=3
	streamLog.Type=2
	streamLog.Open()
	Set wsh=Server.CreateObject("WScript.Shell")
	Do While Not rs.EOF
		source_file=rs(arrDefaultFileListField(filetype))
		If IsNull(source_file) Then
			source_file=""
		Else
			fileExt=LCase(fso.GetExtensionName(source_file))
			oldfilename=source_file
			source_file=Server.MapPath(arrDefaultFileListPath(filetype)&"/"&source_file)
		End If
		If Not fso.FileExists(source_file) Then
			numFailed=numFailed+1
			errMsg=errMsg&vbNewLine&numFailed&"."&rs("STU_NAME")&"的论文《"&rs("THESIS_SUBJECT")&"》没有所需类型的文件。"
		Else
			If Len(arrDefaultFileListNamePostfix(filetype)) Then
				'newfilename=rs("SPECIALITY_NAME")&rs("STU_NAME")&rs("STU_NO")&"-"&arrDefaultFileListNamePostfix(filetype)
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
				'newfilename=rs("SPECIALITY_NAME")&"-"&subject
				newfilename=rs("STU_NAME")&"_"&rs("STU_NO")&"_"&subject
			End If
			newfilename=newfilename&"."&fileExt
			sourcefilelist=sourcefilelist&" """&source_file&""""
			renamefilelist=renamefilelist&" """&oldfilename&""" """&newfilename&""""
			numSucceeded=numSucceeded+1
			
			fso.CopyFile source_file,rarDir&"\"&newfilename
		End If
		numCompleted=numSucceeded+numFailed
		rs.MoveNext()
		If numSucceeded>0 And (numSucceeded Mod 10=0 Or rs.EOF) Then
			streamLog.Flush()
			streamLog.Position=0
			streamLog.WriteText "正在复制文件 "&Round(numCompleted/numRecord,2)*100&"% ("&numCompleted&"/"&numRecord&")……<br/>"
			streamLog.SaveToFile progfile,2
			streamLog.Position=streamLog.Size
		End If
	Loop
	CloseRs rs
	CloseConn conn
	
	' 打包压缩
	cmd=""""&rarExe&""" a -ep -m1 """&rarFile&""" """&rarDir&""""
	Set exec=wsh.Exec(cmd)
	streamLog.WriteText "正在生成压缩文件……<br/>"
	streamLog.SaveToFile progfile,2
	streamLog.Position=streamLog.Size
	exec.StdOut.ReadAll()
	fso.DeleteFolder rarDir
	If numSucceeded=0 Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">所选论文没有<%=arrDefaultFileListName(filetype)%>！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		Response.End()
	End If
	' 添加压缩文件注释
	comment="打包报告："&vbNewLine&numSucceeded&" 个成功，"&numFailed&" 个失败。"&errMsg
	Set stream=fso.CreateTextFile(commentfile)
	stream.Write comment
	stream.Close()
	Set stream=Nothing
	cmd=""""&rarExe&""" c -w"""&tmpDir&""" -z"""&commentfile&""" """&rarFile&""""
	Set exec=wsh.Exec(cmd)
	exec.StdOut.ReadAll()
	Set wsh=Nothing
	streamLog.WriteText "<ok/>导出成功，准备下载……<br/>"&toPlainString(comment)
	streamLog.SaveToFile progfile,2
	streamLog.Close()
	fso.DeleteFile progfile
	fso.DeleteFile commentfile
	Set exec=Nothing
	Set streamLog=Nothing
	Set fso=Nothing
	url="/PaperReview/admin/rar/"&rarFilename
%><script type="text/javascript">
	alert("文件已打包完毕，点击“确定”按钮开始下载。");
	location.href='<%=url%>';
</script><%
End Select
%>
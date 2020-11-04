<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
filetype=Request.Form("filetype")
ids=Request("sel")
numRecord=UBound(Split(ids,","))+1
step=Request.QueryString("step")
Select Case step
Case vbNullString	' 选择页面
	rarFilenamePostfix=Format("（共{0}份）", numRecord)
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>批量下载表格/论文</title>
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body>
<center><font size=4><b>批量下载表格/论文</b><br />
<form action="?step=2" method="POST">
<p>您选择了&nbsp;<%=numRecord%>&nbsp;条论文记录</p>
<p>请选择要下载的文件：<select name="filetype"><option value="0">请选择</option><%
For i=1 To UBound(arrDefaultFileListName)
	If i<>18 Then
%>
<option value="<%=i%>"><%=arrDefaultFileListName(i)%></option><%
	End If
Next
%></select></p>
<p>打包压缩文件名：<input type="text" name="rarfilename" size="40" />.rar&nbsp;</p>
<p><label for="nameAfterId"><input type="checkbox" name="name_after_id" id="nameAfterId" />用“论文序号+论文题目”作为文件名（用于专家评阅）</label></p>
<p><input type="hidden" name="sel" value="<%=ids%>" />
<input type="submit" name="btnsubmit" value="批量下载" />&nbsp;
<input type="button" name="btnclose" value="关 闭" onclick="tabmgr.removeTab(window.index)" /></p></form>
<p align="left"><span id="output" class="output-message"></span></p></center></body>
<script type="text/javascript">
	$(document).ready(function(){
		var progfile=location.origin+"<%=basePath()%>admin/rar/tmp/prog_<%=Session("id")%>.txt";
		$("select").change(function() {
			if(!this.selectedIndex) return;
			$(":text").val(this.options[this.selectedIndex].innerText+"<%=rarFilenamePostfix%>");
		});
		$("form").submit(function() {
			$(":submit").val("正在处理，请稍候...")
				.attr("disabled", true);
			$("#output").html("");
			setTimeout(refreshProgress, 500);
		});
		$(":submit").attr("disabled", false);
		function refreshProgress() {
			$.get(progfile, function (data, status) {
				if(status=="success") {
					$("#output").html(data);
					if(data.match(/<ok\/>/)) {
						$(":submit").val("批量下载").attr("disabled", false);
					} else {
						setTimeout(refreshProgress, 500);
					}
				} else {
					setTimeout(refreshProgress, 500);
				}
			});
		}
	});
</script></html><%
Case 2	' 下载页面

	nameAfterId=Request.Form("name_after_id")="on"
	numRecord=UBound(Split(ids,","))+1
	If numRecord=0 Then
		bError=True
		errMsg="请选择论文！"
	ElseIf Not IsNumeric(filetype) Then
		bError=True
		errMsg="参数无效。"
	Else
		filetype=Int(filetype)
		If filetype<1 Or filetype>UBound(arrDefaultFileListName) Then
			bError=True
			errMsg="请选择要批量下载的文件类型！"
		End If
	End If
	If bError Then
		showErrorPage errMsg,"提示"
	End If
	
	rarFilename=Trim(Request.Form("rarfilename"))
	If LCase(Right(rarFilename,4))=".rar" Then rarFilename=Trim(Left(rarFilename,Len(rarFilename)-4))
	If Len(rarFilename)=0 Then	' 取默认文件名
		rarFilename=Format("{0}（共{1}份）", arrDefaultFileListName(filetype), numRecord)
	End If
	rarFilename=rarFilename&".rar"
	Connect conn
	sql="SELECT * FROM ViewDissertations_admin WHERE ID IN ("&ids&")"
	GetRecordSet conn,rs,sql,count
	If rs.EOF Then
		showErrorPage "所选论文记录不存在！","提示"
	End If
	' 打包文件
	Dim source_file,file_ext,oldfilename,newfilename
	Dim rarExe,rarFile,tmpDir,rarDir,sourcefilelist,renamefilelist,commentfile,progfile
	Dim comment,cmd
	Dim streamLog,wsh
	Dim fso:Set fso=CreateFSO()
	Dim numFailed,numSucceeded,numCompleted
	
	numFailed=0
	numSucceeded=0
	numBatchSize=20
	rarExe=Server.MapPath("rar/Rar.exe")
	rarFile=Server.MapPath(resolvePath("rar",rarFilename))
	tmpDir=Server.MapPath(tempPath())
	rarDir=resolvePath(tmpDir,rarFilename)
	commentfile=resolvePath(tmpDir,"comment_"&Timer&".txt")
	progfile=resolvePath(tmpDir,"prog_"&Session("Id")&".txt")
	
	ensurePathExists tmpDir
	ensurePathExists rarDir
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
			file_ext=LCase(fso.GetExtensionName(source_file))
			oldfilename=source_file
			source_file=Server.MapPath(resolvePath(basePath(),arrDefaultFileListPath(filetype),source_file))
		End If
		If Not fso.FileExists(source_file) Then
			numFailed=numFailed+1
			errMsg=errMsg&vbNewLine&numFailed&"."&rs("STU_NAME")&"的论文《"&rs("THESIS_SUBJECT")&"》没有所需类型的文件。"
		Else
			subject=toFilenameString(rs("THESIS_SUBJECT").Value)
			If nameAfterId Then
				If Len(arrDefaultFileListNamePostfix(filetype)) Then
					newfilename=rs("ID")&"_"&subject&"_"&arrDefaultFileListNamePostfix(filetype)
				Else
					newfilename=rs("ID")&"_"&subject
				End If
			Else
				If Len(arrDefaultFileListNamePostfix(filetype)) Then
					newfilename=rs("STU_NAME")&"_"&rs("STU_NO")&"_"&arrDefaultFileListNamePostfix(filetype)
				Else
					newfilename=rs("STU_NAME")&"_"&rs("STU_NO")&"_"&subject
				End If
			End If
			newfilename=newfilename&"."&file_ext
			sourcefilelist=sourcefilelist&" """&source_file&""""
			renamefilelist=renamefilelist&" """&oldfilename&""" """&newfilename&""""
			numSucceeded=numSucceeded+1
			
			fso.CopyFile source_file,resolvePath(rarDir,newfilename)
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
		Set exec=Nothing
		Set wsh=Nothing
		Set streamLog=Nothing
		Set fso=Nothing
		showErrorPage "所选论文没有"&arrDefaultFileListName(filetype)&"！","提示"
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
	streamLog.WriteText "导出成功，准备下载……<ok/><br/>"&toPlainString(comment)
	streamLog.SaveToFile progfile,2
	streamLog.Close()
	fso.DeleteFile progfile
	fso.DeleteFile commentfile
	Set exec=Nothing
	Set streamLog=Nothing
	Set fso=Nothing
	url=basePath()&"admin/rar/"&rarFilename
%><script type="text/javascript">
	alert("文件已打包完毕，点击“确定”按钮开始下载。");
	setTimeout(function() {
		location.href = "<%=url%>";
		setTimeout(function(){history.go(-1)}, 5000);
	});
</script><%
End Select
%>
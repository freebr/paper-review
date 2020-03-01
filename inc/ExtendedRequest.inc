<%
' 扩展页面请求类
Class ExtendedRequest

	Dim objForm,objFile
	Dim streamReq

	Public Function Form(strForm)
	 strForm=LCase(strForm)
	 If Not objForm.exists(strForm) Then
	   Form=""
	 Else
	   Form=objForm(strForm)
	 End If
	End Function

	Public Function File(strFile)
	 strFile=LCase(strFile)
	 If Not objFile.exists(strFile) Then
	   Set File=New FileInfo
	 Else
	   Set File=objFile(strFile)
	 End If
	End Function

	Private Sub Class_Initialize()
		Dim RequestData,sStart,bCrlf,sInfo,iInfoStart,iInfoEnd,iStart,theFile,charset
		Dim iBytesUploaded,iBlockSize,iFileSize
		Dim sTmpPath,sUploadId,sStreamProgFile,sFilePath,sFileType,sFormValue,sFileName
		Dim iFindStart,iFindEnd
		Dim iFormStart,iFormEnd,sFormName
		Dim isStudent:isStudent=Not IsEmpty(Session("StuId"))
		Dim tStream,streamProg,fso
		If Response.Charset="utf-8" Then
			charset="utf-8"
		Else
			charset="utf-8"
		End If
		iBlockSize=32000	'32kB
		iBytesUploaded=0
		Set objForm=CreateDictionary()
		Set objFile=CreateDictionary()
		If Request.TotalBytes<1 Then Exit Sub
		Set tStream = Server.CreateObject("ADODB.Stream")
		Set streamReq = Server.CreateObject("ADODB.Stream")
		If isStudent Then
		  Set streamProg=Server.CreateObject("ADODB.Stream")
		  Set fso=Server.CreateObject("Scripting.FileSystemObject")
		  sUploadId=Replace(Replace(Request.ServerVariables("SCRIPT_NAME"),"/","_"),".","_")
		  sTmpPath=Server.MapPath("tmp")
		  If Not fso.FolderExists(sTmpPath) Then fso.CreateFolder sTmpPath
		  sStreamProgFile=Server.MapPath("tmp/"&Session("StuId")&sUploadId&".json")
		  streamProg.Type = 2
		  streamProg.Mode = 3
		  streamProg.Charset=charset
		  streamProg.Open
		End If
		streamReq.Type = 1
		streamReq.Mode = 3
		streamReq.Open
		Do
			If iBytesUploaded+iBlockSize>Request.TotalBytes Then
				iBlockSize=Request.TotalBytes-iBytesUploaded
			End If
			streamReq.Write Request.BinaryRead(iBlockSize)
			iBytesUploaded=iBytesUploaded+iBlockSize
			If isStudent Then
		  	streamProg.WriteText "{""bytesUploaded"":"""&iBytesUploaded&""",""bytesTotal"":"""&Request.TotalBytes&"""}"
		  	streamProg.SaveToFile sStreamProgFile,2
		  End If
		Loop While iBytesUploaded<Request.TotalBytes
		If isStudent Then
		  streamProg.Close()
		  fso.DeleteFile sStreamProgFile
		  Set streamProg=Nothing
		  Set fso=Nothing
		End If
		streamReq.Position=0
		RequestData = streamReq.Read
		iFormStart = 1
		iFormEnd = LenB(RequestData)
		bCrlf = ChrB(13) & ChrB(10)
		sStart = MidB(RequestData,1, InStrB(iFormStart,RequestData,bCrlf)-1)
		iStart = LenB(sStart)
		iFormStart=iFormStart+iStart+1
		While (iFormStart + 10) < iFormEnd
			iInfoEnd = InStrB(iFormStart,RequestData,bCrlf & bCrlf)+3
			tStream.Type = 1
			tStream.Mode =3
			tStream.Open()
			streamReq.Position = iFormStart
			streamReq.CopyTo tStream,iInfoEnd-iFormStart
			tStream.Position = 0
			tStream.Type = 2
			tStream.Charset = charset
			sInfo = tStream.ReadText
			tStream.Close
			'取得表单项目名称
			iFormStart = InStrB(iInfoEnd,RequestData,sStart)
			iFindStart = InStr(22,sInfo,"name=""",1)+6
			iFindEnd = InStr(iFindStart,sInfo,"""",1)
			sFormName = LCase(Mid(sinfo,iFindStart,iFindEnd-iFindStart))
			'如果是文件
			If InStr(1,sInfo,"filename=""",1) > 0 Then
				Set theFile=New FileInfo
				'取得文件名
				iFindStart = InStr(iFindEnd,sInfo,"filename=""",1)+10
				iFindEnd = InStr(iFindStart,sInfo,""""&vbCrLf,1)
				sFileName = Mid(sinfo,iFindStart,iFindEnd-iFindStart)
				theFile.FileName=getFileName(sFileName)
				theFile.FilePath=getFilePath(sFileName)
				theFile.FileExt=GetFileExt(sFileName)
				'取得文件类型
				iFindStart = InStr(iFindEnd,sInfo,"Content-Type: ",1)+14
				iFindEnd = InStr(iFindStart,sInfo,vbCr)
				theFile.FileType = Mid(sinfo,iFindStart,iFindEnd-iFindStart)
				theFile.FileStart = iInfoEnd
				theFile.FileSize = iFormStart -iInfoEnd -3
				theFile.FormName=sFormName
				Set theFile.streamReq=streamReq
				If Not objFile.Exists(sFormName) Then
				  objFile.Add sFormName,theFile
				End If
			Else
			'如果是表单项目
				tStream.Type = 1
				tStream.Mode = 3
				tStream.Open()
				streamReq.Position = iInfoEnd
				streamReq.CopyTo tStream,iFormStart-iInfoEnd-3
				tStream.Position = 0
				tStream.Type = 2
				tStream.Charset = charset
			  sFormValue = tStream.ReadText
			  tStream.Close()
				If objForm.Exists(sFormName) Then
				  objForm(sFormName)=objForm(sFormName)&", "&sFormValue
				Else
				  objForm.Add sFormName,sFormValue
				End If
			End If
			iFormStart=iFormStart+iStart+1
		Wend
		RequestData=""
		Set tStream=Nothing
	End Sub

	Private Sub Class_Terminate()
		If Request.TotalBytes>0 Then
			objForm.RemoveAll()
			objFile.RemoveAll()
			Set objForm=Nothing
			Set objFile=Nothing
			streamReq.Close()
			Set streamReq=Nothing
		End If
	End Sub

	Private Function GetFilePath(FullPath)
		If FullPath <> "" Then
		 GetFilePath = Left(FullPath,InStrRev(FullPath, "\"))
		Else
		 GetFilePath = ""
		End If
	End Function

	Private Function GetFileExt(FullPath)
		If FullPath <> "" Then
		 GetFileExt = Mid(FullPath,InStrRev(FullPath, ".")+1)
		Else
		 GetFileExt = ""
		End If
	End Function

	Private Function GetFileName(FullPath)
		If FullPath <> "" Then
		 GetFileName = Mid(FullPath,InStrRev(FullPath, "\")+1)
		Else
		 GetFileName = ""
		End If
	End Function
End Class

Class FileInfo
  Dim FormName,FileName,FilePath,FileSize,FileExt,FileType,FileStart
  Dim streamReq
  Private Sub Class_Initialize
    FileName = ""
    FilePath = ""
    FileSize = 0
    FileStart= 0
    FormName = ""
    FileType = ""
    FileExt  = ""
  End Sub

	Public Function SaveAs(FullPath)
	  Dim dr
	  SaveAs=True
	  If Trim(fullpath)="" Or FileStart=0 Or FileName="" Or Right(fullpath,1)="/" Then exit Function
	  Set dr=CreateObject("ADODB.Stream")
	  dr.Mode=3
	  dr.Type=1
	  dr.Open()
	  streamReq.Position=FileStart
	  streamReq.CopyTo dr,FileSize
	  dr.SaveToFile FullPath,2
	  dr.Close()
	  Set dr=Nothing
	  SaveAs=False
	End Function
End Class
%>

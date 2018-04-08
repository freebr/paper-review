<%Response.Charset="utf-8"%>
<!-- #include File="../inc/db.asp" -->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
filename=Request.QueryString("fn")
If Len(filename)=0 Then
	filename=FormatDateTime(Now(),1)&Int(Timer)
	bFilenameSpec=False
Else
	bFilenameSpec=True
End If
retlink=Request.QueryString("ret")
nTurn=Request.QueryString("turn")
period_id=Request.Form("In_PERIOD_ID")
If Len(period_id)=0 Then period_id=Request.Form("In_PERIOD_ID2")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
class_id=Request.Form("In_CLASS_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
FormGetToSafeRequest(period_id)
If period_id="" Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">信息不完整或格式不正确！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End
End If
If Len(finalFilter) Then PubTerm=" AND "&finalFilter
If Len(teachtype_id) And teachtype_id<>"0" Then
	PubTerm=PubTerm&" AND TEACHTYPE_ID="&toSqlString(teachtype_id)
Else
	teachtype_id="0"
End If
If Len(class_id) And class_id<>"0" Then
	PubTerm=PubTerm&" AND SPECIALITY_ID="&toSqlString(class_id)
Else
	class_id="0"
End If
If Len(enter_year) And enter_year<>"0" Then
	PubTerm=PubTerm&" AND ENTER_YEAR="&toSqlString(enter_year)
Else
	enter_year="0"
End If
If Len(query_task_progress) And query_task_progress<>"-1" Then
	PubTerm=PubTerm&" AND TASK_PROGRESS="&toSqlString(query_task_progress)
End If
If Len(query_review_status) And query_review_status<>"-1" Then
	PubTerm=PubTerm&" AND REVIEW_STATUS="&toSqlString(query_review_status)
End If

Class ExcelGen
	Private spSheet
	Private iColOffset
	Private iRowOffset
	
	Sub Class_Initialize()
	  Set spSheet=Server.CreateObject("OWC11.Spreadsheet")
	  spSheet.DisplayToolBar=True
	  iRowOffSet=2
	  iColOffSet=2
	End Sub
	
	Sub Class_Terminate()
	  Set spSheet=Nothing 'Clean up
	End Sub
	
	Public Property Let ColumnOffset(iColOff)
	  If iColOff > 0 then
	    iColOffSet=iColOff
	  Else
	    iColOffSet=2
	  End If
	End Property
	
	Public Property Let RowOffset(iRowOff)
	  If iRowOff > 0 then
	     iRowOffSet=iRowOff
	  Else
	     iRowOffSet=2
	  End If
	End Property
	
	Function GenerateWorksheet(arrFields,arrRs,arrSheetName)
	  'Populates the Excel worksheet based on a Recordset's contents
	  'Start by displaying the titles
	  Dim iCol,iRow,colId
	  Dim parentFieldName,colSubfieldBegin,colSubfieldEnd,bNewParentField
	  Dim fieldSizeDef
	  Dim i,j,tmp,cellid,arr
	  Dim nSheetId,nRecNum,sheet
	  
	  For nSheetId=0 To UBound(arrRs)
		  nRecNum=0
	  	If nSheetId>0 Then
	  		Set sheet=sheet.Next
	  		If sheet Is Nothing Then
	  			Set sheet=spSheet.Sheets.Add()
	  		End If
	  		sheet.Activate()
	  	Else
	  		Set sheet=spSheet.ActiveSheet
	  	End If
	  	sheet.Name=arrSheetName(nSheetId)
	  	If Not arrRs(nSheetId).EOF Then
				iCol=iColOffset
				iRow=iRowOffset
				parentFieldName=""
				ReDim fieldSizeDef(UBound(arrFields(nSheetId)),1)
				For i=0 To UBound(arrFields(nSheetId))
					bNewParentField=False
					colId=Chr(i+65)
					If Len(arrFields(nSheetId)(i))>0 Then
						arr=Split(arrFields(nSheetId)(i),"*")
						strFieldName=arr(0)
						If UBound(arr)=2 Then	' 指定了字段大小
							fieldSizeDef(i,0)=arr(1)
							fieldSizeDef(i,1)=arr(2)
						End If
					End If
					j=InStr(strFieldName,".")
					If j Then
						tmp=Left(strFieldName,j-1)
						strFieldName=Mid(strFieldName,j+1)
						If tmp<>parentFieldName Then
							bNewParentField=True
						ElseIf i=UBound(arrFields(nSheetId)) Then
							bNewParentField=True
							colSubfieldEnd=colId
						End If
					Else
						bNewParentField=True
						tmp=""
					End If
					If bNewParentField Then
						If Len(parentFieldName) Then
							cellid=colSubfieldBegin&iRowOffset
							spSheet.Range(cellid).Value=parentFieldName
							spSheet.Range(cellid).Interior.ColorIndex=37
							spSheet.Range(cellid).Font.Bold=True
							spSheet.Range(cellid).Font.Size=11
							spSheet.Range(colSubfieldBegin&iRowOffset&":"&colSubfieldEnd&iRowOffset).Merge()
							spSheet.Range(colSubfieldBegin&iRowOffset&":"&colSubfieldEnd&iRowOffset).HorizontalAlignment=-4108
						End If
						If Len(tmp)=0 Then
							iRow=iRowOffset
							spSheet.Range(colId&iRowOffset&":"&colId&(iRowOffset+1)).Merge()
						End If
						parentFieldName=tmp
						colSubfieldBegin=colId
					End If
					colSubfieldEnd=colId
					If Len(parentFieldName) Then
						iRow=iRowOffset+1
					End If
					spSheet.Cells(iRow, iCol).Value=strFieldName
					spSheet.Cells(iRow, iCol).Interior.ColorIndex=37
					spSheet.Cells(iRow, iCol).Font.Bold=True
					spSheet.Cells(iRow, iCol).Font.Size=11
					spSheet.Cells(iRow, iCol).HorizontalAlignment=-4108 ' 居中
					If Len(fieldSizeDef(i,0))<>0 Then
						spSheet.Cells(iRow, iCol).ColumnWidth=fieldSizeDef(i,0)
					ElseIf Len(strFieldName)=0 then
						spSheet.Cells(iRow, iCol).ColumnWidth=0
					Else
						spSheet.Columns(iCol).AutoFit()
					End If
					iCol=iCol+1
				Next
				'Display all of the data
				iRow=iRowOffset+1
				Do While Not arrRs(nSheetId).EOF
				 	iRow=iRow + 1
				 	iCol=iColOffset
				 	For j=0 To UBound(arrFields(nSheetId))
				    If IsNull(arrRs(nSheetId)(j)) Then
				      spSheet.Cells(iRow, iCol).Value=""
				    Else
				      spSheet.Cells(iRow, iCol).Value=CStr(arrRs(nSheetId)(j))
							spSheet.Cells(iRow, iCol).Font.Bold=False
							spSheet.Cells(iRow, iCol).Font.Italic=False
							spSheet.Cells(iRow, iCol).Font.Size=10
							If Len(fieldSizeDef(j,0))=0 Then
								spSheet.Columns(iCol).AutoFit()
							End If
							If Len(fieldSizeDef(j,1))<>0 Then
								spSheet.Cells(iRow, iCol).RowHeight=fieldSizeDef(j,1)
							End If
				    End If
				  	iCol=iCol + 1
					Next
					nRecNum=nRecNum+1
					arrRs(nSheetId).MoveNext()
				Loop
			End If
		Next
		spSheet.Sheets(2).Activate()
		GenerateWorksheet=nRecNum
	End Function
	
	Function SaveWorksheet(strFileName)
		'Save the worksheet to a specified filename
		On Error Resume Next
		Call spSheet.Export(strFileName, 0)
		SaveWorksheet=(Err.Number=0)
	End Function
End Class

Dim arrFields,rs(1),arrSheetName
arrFields=Array(Array("","学位类别","专业名称","总数",_
					"送审结果.同意答辩","送审结果.适当修改","送审结果.重大修改","送审结果.加送两份","送审结果.延期送审","送审结果.未齐",_
					"总体评价.优","总体评价.良","总体评价.中","总体评价.差","导师审核.同意","导师审核.不同意"),_
								Array("状态","论文题目*37*80","作者姓名","学号","专业","研究方向","论文形式","导师","开题报告","中期检查表","预答辩意见书","答辩审批材料","复制比","专家一姓名","专家一工作单位","专家二姓名","专家二工作单位","送审结果1","送审结果2","处理意见","答辩修改意见*55*80","答辩成绩","分会修改意见*55*80"))
If nTurn<>0 Then turnPostfix="(第 "&nTurn&" 批)"
arrSheetName=Array("送审结果统计表"&turnPostfix,"全部论文列表"&turnPostfix)

Connect conn
selectFields="dbo.getThesisStatusText(1,TASK_PROGRESS,1)+'，'+dbo.getThesisStatusText(2,REVIEW_STATUS,1),THESIS_SUBJECT,STU_NAME,(''''+STU_NO) AS STU_NO,SPECIALITY_NAME,RESEARCHWAY_NAME,THESIS_FORM,TUTOR_NAME,"&_
						 "dbo.getStatusOfReviewFile(ID,0,0),dbo.getStatusOfReviewFile(ID,0,1),dbo.getStatusOfReviewFile(ID,0,2),dbo.getStatusOfReviewFile(ID,0,3),"&_
						 "dbo.getDetectResultString(ID) AS RATIO,EXPERT_NAME1,EXPERT_WORKPLACE1,EXPERT_NAME2,EXPERT_WORKPLACE2,dbo.getReviewResultText(LEFT(REVIEW_RESULT,1)) AS REVIEW_RESULT1,dbo.getReviewResultText(SUBSTRING(REVIEW_RESULT,3,1)) AS REVIEW_RESULT2,dbo.getFinalResultText(RIGHT(REVIEW_RESULT,1)) AS FINAL_RESULT,DEFENCE_EVAL,dbo.getDefenceResultText(DEFENCE_RESULT),INSTRUCT_MODIFY_EVAL"
If nTurn=0 Then
	' 导出送审结果统计表
	Set rs(0)=conn.Execute("EXEC getTestThesisReviewStatsList "&period_id&",0")
	' 导出送审论文列表
	sql="SELECT "&selectFields&" FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE PERIOD_ID="&period_id&PubTerm
	Set rs(1)=conn.Execute(sql)
Else	' 按批次导出
	exportFilter="PERIOD_ID="&period_id&PubTerm&" AND NOT EXIST(SELECT STU_ID FROM EXPORT_INFO WHERE PERIOD_ID="&period_id&" AND STU_ID=A.STU_ID)"
	Set rs(0)=conn.Execute("EXEC getTestThesisReviewStatsList "&period_id&",1")
	sql="SELECT "&selectFields&" FROM VIEW_TEST_THESIS_REVIEW_INFO A WHERE "&exportFilter
	Set rs(1)=conn.Execute(sql)
	sql="INSERT INTO EXPORT_INFO (STU_ID,PERIOD_ID,TURN_ID) SELECT STU_ID,"&period_id&","&nTurn&" FROM TEST_THESIS_REVIEW_INFO A WHERE "&exportFilter
	conn.Execute sql
End If

Dim fso
Set fso=Server.CreateObject("Scripting.FileSystemObject")
exportBaseDir=Server.MapPath("export")
exportSpecDir=exportBaseDir&"/spec"
If Not fso.FolderExists(exportBaseDir) Then
	fso.CreateFolder(exportBaseDir)
End If
If Not fso.FolderExists(exportSpecDir) Then
	fso.CreateFolder(exportSpecDir)
End If
Do
	s="/"&filename&".xls"
	If bFilenameSpec Then
		excelPath=exportSpecDir&s
		exportPath="export/spec"&s
	Else
		excelPath=exportBaseDir&s
		exportPath="export"&s
	End If
	filename=filename&"(2)"
Loop While fso.FileExists(excelPath)
Set fso=Nothing

Set objExcel=New ExcelGen
objExcel.RowOffSet=1
objExcel.ColumnOffSet=1
nRecNum=objExcel.GenerateWorksheet(arrFields,rs,arrSheetName)
If nRecNum>0 Then
	If objExcel.SaveWorksheet(excelPath) then
		nResult=1
	Else
		nResult=2
	End If
Else
	nResult=0
End If
Set objExcel=Nothing
CloseRs rs(0)
CloseRs rs(1)
CloseConn conn
If Len(retlink) Then
	Response.Redirect retlink
End If
%><html><head><link href="../css/admin.css" rel="stylesheet" type="text/css" /></head><body bgcolor="ghostwhite"><p align="center"><%
Select Case nResult
Case 0
%>未生成Excel文件，因为没有数据库记录!<%
Case 1
%>已保存为Excel文件，<a href="<%=exportPath%>" target="_blank">点击下载</a>（直接点击打开，点击右键另存为下载）<%
Case 2
%>保存过程中发生错误!<%
End Select
%><br><a href="javascript:history.go(-1)">返回</a></p></body></html>
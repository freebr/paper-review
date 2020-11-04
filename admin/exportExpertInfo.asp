<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")

ids=Request.Form("sel")
finalFilter=Request.Form("finalFilter2")
Dim PubTerm:PubTerm=""
If Not IsEmpty(ids) Then PubTerm=PubTerm&" AND TEACHER_ID IN ("&ids&")"
If Len(finalFilter) Then PubTerm=PubTerm&" AND "&finalFilter

Class ExcelGen
	Private spSheet
	Private iColOffset
	Private iRowOffset
	
	Sub Class_Initialize()
	  Set spSheet = Server.CreateObject("OWC11.Spreadsheet")
	  spSheet.DisplayToolBar = True
	  iRowOffSet = 2
	  iColOffSet = 2
	End Sub
	
	Sub Class_Terminate()
	  Set spSheet = Nothing 'Clean up
	End Sub
	
	Public Property Let ColumnOffset(iColOff)
	  If iColOff > 0 then
	    iColOffSet = iColOff
	  Else
	    iColOffSet = 2
	  End If
	End Property
	
	Public Property Let RowOffset(iRowOff)
	  If iRowOff > 0 then
	     iRowOffSet = iRowOff
	  Else
	     iRowOffSet = 2
	  End If
	End Property
	
	Function GenerateWorksheet(arrFields,rs,sheetName)
	  'Populates the Excel worksheet based on a Recordset's contents
	  'Start by displaying the titles
	  Dim iCol,iRow,i
	  Dim nRecNum,sheet
	  
	  nRecNum=0
	 	Set sheet=spSheet.ActiveSheet
	  If Not rs.EOF Then
			iCol=iColOffset
			iRow=iRowOffset
  		sheet.Name=sheetName
			For i=0 To UBound(arrFields)
				strFieldName = arrFields(i)
				spSheet.Cells(iRow, iCol).Value = strFieldName
				spSheet.Cells(iRow, iCol).Font.Bold = True 
				spSheet.Cells(iRow, iCol).Font.Italic = False
				spSheet.Cells(iRow, iCol).Font.Size = 10
				spSheet.Cells(iRow, iCol).HorizontalAlignment = -4108 ' 居中
				spSheet.Columns(iCol).AutoFit
				iCol = iCol + 1
			Next
			'Display all of the data
			Do While Not rs.EOF
			 	iRow = iRow + 1
			 	iCol = iColOffset
			 	For j=0 To UBound(arrFields)
			    If IsNull(rs(j)) Then
			      spSheet.Cells(iRow, iCol).Value = ""
			    Else
			    	If iCol=iColOffset Then
			    		spSheet.Cells(iRow, iCol).Value = iRow-iRowOffset
			    	Else
			      	spSheet.Cells(iRow, iCol).Value = "'"&CStr(rs(j))
			      End If
						spSheet.Cells(iRow, iCol).Font.Bold = False
						spSheet.Cells(iRow, iCol).Font.Italic = False
						spSheet.Cells(iRow, iCol).Font.Size = 10 
						spSheet.Columns(iCol).AutoFit()
			    End If
			  	iCol = iCol + 1
				Next
				nRecNum=nRecNum+1
				rs.MoveNext()
			Loop
		End If
		GenerateWorksheet=nRecNum
	End Function
	
	Function SaveWorksheet(strFileName)
		'Save the worksheet to a specified filename
		On Error Resume Next
		Call spSheet.Export(strFileName, 0)
		SaveWorksheet = (Err.Number = 0)
	End Function
End Class

Dim arrFields,rs
arrFields=Array("序号","姓名","登录名","密码","职称","学科专长","单位名称（含院系）","通信地址","邮编","联系电话（办公室）","联系电话（移动）","电子邮箱","身份证号码","银行账户","开户行")

Connect conn
' 导出评阅专家名单
sql="SELECT ID,EXPERT_NAME,TEACHERNO,PASSWORD,PRO_DUTY_NAME,EXPERTISE,WORKPLACE,ADDRESS,MAILCODE,TELEPHONE,MOBILE,EMAIL,IDCARD_NO,BANK_ACCOUNT,BANK_NAME,TEACHER_ID FROM ViewExpertInfo WHERE VALID=1 "&PubTerm
Set rs=conn.Execute(sql)

Dim fso
Set fso=CreateFSO()
exportBaseDir=Server.MapPath("export")
If Not fso.FolderExists(exportBaseDir) Then
	fso.CreateFolder(exportBaseDir)
End If
filename=timestamp()
s="/"&filename&".xls"
excelPath=exportBaseDir&s
exportPath="export"&s
Set fso=Nothing

Set objExcel=New ExcelGen
objExcel.RowOffSet=1
objExcel.ColumnOffSet=1
nRecNum=objExcel.GenerateWorksheet(arrFields,rs,"评阅专家信息表")
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
CloseRs rs
CloseConn conn
%><html><head><% useStylesheet "admin" %></head><body><p align="center"><%
Select Case nResult
Case 0
%>未生成Excel文件，因为没有数据库记录!<%
Case 1
%>已保存为Excel文件，<a href="<%=exportPath%>" target="_blank">点击下载</a>（直接点击打开，点击右键另存为下载）<%
Case 2
%>保存过程中发生错误!<%
End Select
%><br><a href="javascript:history.go(-1)">返回</a></p></body></html>
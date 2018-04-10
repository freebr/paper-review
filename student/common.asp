<!--#include file="../inc/global.inc"-->
<%
Function toYearMonth(ByVal year,ByVal month)
	' 返回形如 yyyy.mm 的日期格式
	If Len(month)=1 Then month="0"&month
	toYearMonth=year&"."&month
End Function
Function addFormInfoToArray(tg,arr,fieldName,propName)
	ReDim arr(Request.Form(fieldName).Count-1)
	For i=0 To UBound(arr)
		arr(i)=Request.Form(fieldName)(i+1)
	Next
	tg.addInfo propName,arr
	addFormInfoToArray=1
End Function
Function isMatched(pattern,s)
	' 判断指定字符串是否满足指定模式
	Dim regEx:Set regEx=New RegExp
	regEx.Pattern=pattern
	isMatched=regEx.Test(s)
	Set regEx=Nothing
End Function
Function loadResearchwayList(stu_type)
	Dim arr
	If stu_type=5 Then ' 工程硕士
		arr=Array("--【工业工程领域研究方向：】","精益管理","质量管理","工作研究与人因工程","物流与供应链管理","服务运营与技术创新","信息管理与电子商务","系统分析方法与优化技术","工业工程相关方向",_
		"--【项目管理领域研究方向：】","项目计划与控制","项目质量管理","项目可行性与评估","项目投融资与风险管理","项目管理信息化","企业项目化管理","项目管理其他相关方向",_
		"--【物流工程领域研究方向：】","物流系统规划","电子商务物流","仓储与库存管理","国际物流管理","供应链协调与整合","采购与供应管理","供应链金融","物流工程其他相关方向")
	ElseIf stu_type=6 Or stu_type=7 Then	' MBA/EMBA
		arr=Array("创新与创业","项目管理与工业工程","组织与人力资源","财务与金融","服务与营销","企业战略")
	Else
		arr=Array("")
	End If	
	loadResearchwayList=arr
End Function
Function getTeacherIdByName(name)
	If IsNull(name) Then
		getTeacherIdByName=-1
		Exit Function
	End If
	Dim conn,rsTeacher,sql,num
	name=Replace(name," ",vbNullString)
	name=Replace(name,"　",vbNullString)
	name=Replace(name,"'","''")
	name=Replace(name,"""","""""")
	Connect conn
	sql="SELECT TEACHERID,TEACHERNAME FROM TEACHER_INFO WHERE TEACHERNAME='"&name&"' AND VALID=0"
	GetRecordSetNoLock conn,rsTeacher,sql,num
	If rsTeacher.EOF Then
		getTeacherIdByName=-1
	Else
		getTeacherIdByName=rsTeacher("TEACHERID")
	End If
	CloseRs rsTeacher
	CloseConn conn
End Function
Function getTeachTypeNameById(teachtype_id)
	Dim ret
	Select Case UCase(teachtype_id)
	Case 5:ret="工程硕士"
	Case 6:ret="MBA"
	Case 7:ret="EMBA"
	Case 9:ret="MPAcc"
	Case Else:ret="未知类别"
	End Select
	getTeachTypeNameById=ret
End Function
Function getProDutyNameOf(tid)
	Dim conn,rs,sql,num
	Connect conn
	sql="SELECT PRO_DUTYNAME FROM VIEW_TEACHER_INFO WHERE TEACHERID="&tid
	GetRecordSetNoLock conn,rs,sql,num
	If Not rs.EOF Then
		getProDutyNameOf=rs(0)
	End If
	CloseRs rs
	CloseConn conn
End Function
Function getReviewResult(n)
	Dim ret
	Select Case n
	Case 1:ret="同意答辩"
	Case 2:ret="需做适当修改"
	Case 3:ret="需做重大修改"
	Case 4:ret="不同意答辩"
	Case 5:ret="尚未返回"
	End Select
	getReviewResult=ret
End Function
Function getFinalResult(n)
	Dim ret
	Select Case n
	Case 1:ret="同意答辩"
	Case 2:ret="适当修改"
	Case 3:ret="重大修改"
	Case 4:ret="加送两份"
	Case 5:ret="延期送审"
	Case 6:ret="暂无"
	End Select
	getFinalResult=ret
End Function
Function getDefenceResult(n)	' 按代码返回相应的答辩成绩
	Dim arr:arr=Array("未录入","优秀","良好","及格","不及格")
	getDefenceResult=arr(n)
End Function
Function sendEmailToTutor(filename)
	Dim conn,rs,sql,num
	Dim arrSemInfo,arrMailId
	Dim stuname,stuno,stuclass,stuspec,stumail,subject,tutorname,tutormail,uploadtime,fieldval,bSuccess,logtxt
	arrSemInfo=getCurrentSemester()
	arrMailId=getThesisReviewSystemMailIdByType(Now)
	Connect conn
	sql="SELECT * FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE STU_ID="&Session("StuId")&" AND PERIOD_ID="&arrSemInfo(3)
	GetRecordSetNoLock conn,rs,sql,num
	If rs.EOF Then
		CloseRs rs
		CloseConn conn
		sendEmailToTutor=0
		Exit Function
	End If
	stuname=rs("STU_NAME")
	stuno=rs("STU_NO")
	stuclass=rs("CLASS_NAME")
	stuspec=rs("SPECIALITY_NAME")
	stumail=rs("EMAIL")
	subject=rs("THESIS_SUBJECT")
	tutorname=rs("TUTOR_NAME")
	tutormail=rs("TUTOR_EMAIL")
	uploadtime=Now()
	uploadtime=FormatDateTime(uploadtime,1)&" "&FormatDateTime(uploadtime,4)
	fieldval=Array(stuname,stuno,stuclass,stuspec,stumail,subject,tutorname,tutormail,filename,uploadtime)
	bSuccess=sendAnnouncementEmail(arrMailId(5),tutormail,fieldval)
	logtxt="学生["&Session("Stuname")&"]在论文电子评阅系统执行上传["&filename&"]操作，发送邮件给导师["&tutorname&":"&tutormail&"]"
	If bSuccess Then
		logtxt=logtxt&"成功。"
	Else
		logtxt=logtxt&"失败。"
	End If
	WriteLog logtxt
	CloseRs rs
	CloseConn conn
	sendEmailToTutor=1
End Function
Function reviewResultList(ctlname,sel,showtip)	' 显示评审结果选择框
	Dim arr,i
	arr=Array("","A","B","C","D","E")
%><select name="<%=ctlname%>"><%
	If showtip Then %><option value="0">请选择</option><% End If
	For i=1 To UBound(arr)
%><option value="<%=i%>"<% If sel=i Then Response.Write " selected"%>><%=arr(i)%></option><%
	Next %>
</select><%
End Function
Function finalResultList(ctlname,sel,showtip)	' 显示处理意见选择框
	Dim arr,i
	arr=Array("","I","II","III","IV","V","VI")
%><div onmousedown="return false" onkeydown="return false"><select name="<%=ctlname%>"><%
	If showtip Then %><option value="0">暂无</option><% End If
	For i=1 To UBound(arr)
%><option value="<%=i%>"<% If sel=i Then Response.Write " selected"%>><%=arr(i)%></option><%
	Next %>
</select></div><%
End Function
Function getClientInfo(cli)
	Dim conn,rs,sql,result,i
	Dim sem_info:sem_info=getCurrentSemester()
	Connect conn
	sql="SELECT STU_CLIENT_STATUS,STU_TABLE1_STARTDATE,STU_TABLE1_ENDDATE,STU_TABLE2_STARTDATE,STU_TABLE2_ENDDATE,STU_TABLE3_STARTDATE,STU_TABLE3_ENDDATE,STU_TABLE4_STARTDATE,STU_TABLE4_ENDDATE,STU_DETECT_STARTDATE,STU_DETECT_ENDDATE,STU_REVIEW_STARTDATE,STU_REVIEW_ENDDATE,STU_MODIFY_STARTDATE,STU_MODIFY_ENDDATE,STU_FINAL_STARTDATE,STU_FINAL_ENDDATE FROM TEST_THESIS_REVIEW_SYSTEM WHERE USE_YEAR="&sem_info(0)&" AND USE_SEMESTER="&sem_info(1)&" AND VALID=1"
	GetRecordSetNoLock conn,rs,sql,result
	If rs.EOF Then
		stuclient.SystemStatus=STUCLI_STATUS_CLOSED
	Else
		stuclient.SystemStatus=STUCLI_STATUS_OPEN
		cli.setClientStatus rs(0).Value
		For i=STUCLI_OPR_TABLE1 To STUCLI_OPR_FINAL
			cli.setOpentime i,STUCLI_OPENTIME_START,rs(2*i-1).Value
			cli.setOpentime i,STUCLI_OPENTIME_END,rs(2*i).Value
		Next
	End If
	CloseRs rs
	CloseConn conn
	getClientInfo=1
End Function
If 0 And Not hasPrivilege(Session("writeprivileges"),"SA8") And Not hasPrivilege(Session("readprivileges"),"SA8") Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">您没有权限！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End
End If
Dim stuclient:Set stuclient=New StudentClientInfo
getClientInfo(stuclient)
If stuclient.SystemStatus=STUCLI_STATUS_CLOSED Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">电子评阅系统未启用！</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End
End If
Dim arrTable,arrTblThesis,arrTableStatText,arrStuOprName,arrStep
arrTable=Array("","开题报告表","中期检查表","预答辩申请表","答辩及授予学位审批材料")
arrTblThesis=Array("","开题论文","中期论文","预答辩论文")
arrTblThesisDetail=Array("","开题论文（已完成的论文部分）","中期论文","预答辩论文")
arrTableStatText=Array("—","待审核","审核不通过","审核通过")
arrStuOprName=Array("","开题报告表","中期检查表","预答辩申请表","答辩及授予学位审批材料","送检论文和送审论文","送审论文","答辩论文","定稿论文")
arrStep=Array("","提交送检和送审论文","导师不同意检测","导师同意检测","论文一次检测未通过","论文二次检测未通过","论文二次检测已通过，等候导师同意送审","导师不同意送审","论文检测已通过，导师同意送审","专家正在评阅","专家完成评阅","导师确认评阅结果","提交答辩论文","答辩论文未通过","答辩论文已通过","答辩委员会给出修改意见","教指会分会给出修改意见","提交定稿论文")
%>
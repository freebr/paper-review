<%
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
	sql="SELECT PRO_DUTYNAME FROM ViewTeacherInfo WHERE TEACHERID="&tid
	GetRecordSetNoLock conn,rs,sql,num
	If Not rs.EOF Then
		getProDutyNameOf=rs(0)
	End If
	CloseRs rs
	CloseConn conn
End Function

Function getReviewResultText(n)
	Dim ret
	Select Case n
	Case 1:ret="同意答辩"
	Case 2:ret="需做适当修改"
	Case 3:ret="需做重大修改"
	Case 4:ret="不同意答辩"
	Case 5:ret="尚未返回"
	End Select
	getReviewResultText=ret
End Function

Function getFinalResultText(n)
	Dim ret
	Select Case n
	Case 1:ret="同意答辩"
	Case 2:ret="适当修改"
	Case 3:ret="重大修改"
	Case 4:ret="加送两份"
	Case 5:ret="延期送审"
	Case 6:ret="暂无"
	End Select
	getFinalResultText=ret
End Function

Function getSectionAccessibilityInfo(activity_id, stu_type_id, section_id, dissertation_status)
	Dim section, time_flag, tip
	Dim accessible:accessible=False
	If Not isActivityOpen(activity_id) Then
		Set section=getSectionInfo(Null, Null, section_id)
		time_flag=-3
	Else
		Set section=getSectionInfo(activity_id, stu_type_id, section_id)
		time_flag=compareNowWithSectionTime(section)
		accessible=time_flag=0
	End If
	If Not accessible Then
		If time_flag=-3 Then
			tip=Format("当前评阅活动已关闭，不能执行【{0}】操作。", section("Name"))
		ElseIf time_flag=-2 Then
			tip=Format("【{0}】环节已关闭，不能执行操作。", section("Name"))
		ElseIf time_flag<>0 Then
			tip=Format("【{0}】环节开放时间为{1}至{2}，当前不在开放时间内，不能执行操作。",_
				section("Name"), toDateTime(section("StartTime"), 1), toDateTime(section("EndTime"), 1))
		Else
			tip=Format("当前状态为【{0}】，不能执行操作。", dissertation_status)
		End If
	End If
	Dim dict:Set dict=CreateDictionary()
	dict.Add "section", section
	dict.Add "time_flag", time_flag
	dict.Add "accessible", accessible
	dict.Add "tip", tip
	Set getSectionAccessibilityInfo=dict
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

Function getPinyinOfName(name)
	getPinyinOfName=(New PinyinQuery).getNamePinyinOf(name)
End Function

If Not hasPrivilege(Session("writeprivileges"),"SA8") And Not hasPrivilege(Session("readprivileges"),"SA8") Then
	showErrorPage "您没有访问本系统的权限！", "提示"
End If

Dim arrTable:arrTable=Array("","开题报告表","中期检查表","预答辩申请表","答辩及授予学位审批材料")
Dim arrTblThesis:arrTblThesis=Array("","开题论文","中期论文","预答辩论文")
Dim arrTblThesisDetail:arrTblThesisDetail=Array("","开题论文（已完成的论文部分）","中期论文","预答辩论文")
Dim arrTableStatText:arrTableStatText=Array("—","待审核","审核不通过","审核通过")
Dim arrStuOprName:arrStuOprName=Array("","开题报告表","中期检查表","预答辩申请表","答辩及授予学位审批材料","送检论文和送审论文","送审论文","答辩论文","定稿论文")
Dim arrStep:arrStep=Array("","提交送检和送审论文","导师不同意检测","导师同意检测","论文一次检测未通过","论文二次检测未通过","论文二次检测已通过，等候导师同意送审","导师不同意送审","论文检测已通过，导师同意送审","专家正在评阅","专家完成评阅","导师确认评阅结果","提交答辩论文","答辩论文未通过","答辩论文已通过","答辩委员会给出修改意见","教指会分会给出修改意见","提交定稿论文")
Dim arrFileListName:arrFileListName=Array("","开题报告表","开题论文","中期检查表","中期论文","预答辩申请表","预答辩论文","答辩及授予学位审批材料","一次送检论文","二次送检论文","送审论文","答辩论文","教指委盲评论文","定稿论文","一次送检论文检测报告","二次送检论文检测报告","论文评阅书 1","论文评阅书 2")
Dim arrFileListNamePostfix:arrFileListNamePostfix=Array("","开题报告表","开题论文","中期检查表","中期论文","预答辩申请表","预答辩论文","答辩审批材料","","","","","","","一次检测报告","二次检测报告","论文评阅书(1)","论文评阅书(2)")
Dim arrFileListPath:arrFileListPath=Array("","/PaperReview/student/upload","/PaperReview/student/upload","/PaperReview/student/upload","/PaperReview/student/upload","/PaperReview/student/upload","/PaperReview/student/upload","/PaperReview/student/upload","/PaperReview/student/upload","/PaperReview/student/upload","/PaperReview/student/upload","/PaperReview/student/upload","/PaperReview/student/upload","/PaperReview/student/upload","/PaperReview/admin/upload/report","/PaperReview/admin/upload/report","/PaperReview/expert/export","/PaperReview/expert/export")
Dim arrFileListField:arrFileListField=Array("","TABLE_FILE1","TBL_THESIS_FILE1","TABLE_FILE2","TBL_THESIS_FILE2","TABLE_FILE3","TBL_THESIS_FILE3","TABLE_FILE4","DETECT_THESIS1","DETECT_THESIS2","THESIS_FILE2","THESIS_FILE3","THESIS_FILE4","THESIS_FILE5","DETECT_REPORT1","DETECT_REPORT2","ReviewFile1","ReviewFile2")

Dim arrEthnic:arrEthnic=Array(_
"汉族","阿昌族","白族","保安族","布朗族","布依族","朝鲜族","达斡尔族","傣族","德昂族","侗族","东乡族","独龙族","鄂伦春族","俄罗斯族","鄂温克族","高山族","仡佬族","哈尼族","哈萨克族","赫哲族","回族","基诺族","京族","景颇族","柯尔克孜族","拉祜族","黎族","傈僳族","珞巴族","满族","毛南族","门巴族","蒙古族","苗族","仫佬族","纳西族","怒族","普米族","羌族","撒拉族","畲族","水族","塔吉克族","塔塔尔族","土族","土家族","佤族","维吾尔族","乌兹别克族","锡伯族","瑶族","彝族","裕固族","藏族","壮族")

Dim arrPoliticalStatus:arrPoliticalStatus=Array(_
"中共党员","中共预备党员","共青团员","民革会员","民盟盟员","民建会员","民进会员","农公党党员","致公党党员","九三学社社员","台盟盟员","无党派人士","群众")

Dim arrResearchField:arrResearchField=Array(_
"物流工程","项目管理","工业工程")

Dim arrIssueSource:arrIssueSource=Array(_
"02.973、863项目","04.国家社科规划、基金项目","05.教育部人文、社会科学研究项目","06.国家自然科学基金项目","07.中央、国家各部门项目","09.省（自治区、直辖市）项目","12.国际合作研究项目","13.与港、澳、台合作研究项目","14.企、事业单位委托项目","15.外资项目","16.学校自选项目","17.国防项目","90.非立项","99.其它项目")

Dim arrDissertationType:arrDissertationType=Array(_
"2.应用研究","4.其它")

Dim arrResearchFieldEn:arrResearchFieldEn=Array(_
"Logistics Engineering","Project Management","Industrial Engineering")

Dim arrIssueSourceEn:arrIssueSourceEn=Array(_
"02.Projects sponsored by the State Key Development Program for Basic Research of China, projects sponsored by the State Key Development Program for Basic Research of China",_
"04.Projects sponsored by the State Social Science Fund of China",_
"05.Projects sponsored by the Ministry of Education on humanities and social science",_
"06.Projects of cooperation with Hong Kong, Macao and Taiwan",_
"07.Projects sponsored by other ministries of China's State Council",_
"09.Projects sponsored by provincial governments",_
"12.Projects of international cooperation",_
"13.Projects of cooperation with Hong Kong, Macao and Taiwan",_
"14.Projects sponsored by enterprises",_
"15.Projects sponsored by the foreign investment",_
"16.Projects sponsored by universities",_
"17.Projects sponsored by the National Defense",_
"90.Non-established projects",_
"99.Other projects")

Dim arrDissertationTypeEn:arrDissertationTypeEn=Array(_
"2.Applied research","4.Others")
%>
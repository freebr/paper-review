<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<!--#include file="../inc/setEditor.asp"-->
<!--#include file="../inc/ckeditor/ckeditor.asp"-->
<!--#include file="../inc/ckfinder/ckfinder.asp"-->
<!--#include file="common.asp"-->
<%Response.Expires=-1
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
curstep=Request.QueryString("step")
If Len(curstep)=0 Or Not IsNumeric(curstep) Then curstep="1"
sem_info=getCurrentSemester()
Dim arrMailId,arrMailSubject,arrStuOprFlag,arrStuOprName
Dim tutor_startdate,tutor_enddate,exp_startdate,exp_enddate
Dim stu_startdate():ReDim stu_startdate(OPRTYPE_COUNT)
Dim stu_enddate():ReDim stu_enddate(OPRTYPE_COUNT)
Dim stu_clientstatus
arrStuOprFlag=Array("","TABLE1","TABLE2","TABLE3","TABLE4","DETECT","REVIEW","MODIFY","FINAL")
arrStuOprName=Array("","开题报告表","中期检查表","预答辩申请表","答辩审批材料","送检论文","送审论文","答辩论文","定稿论文")
arrMailSubject=Array("","论文送审通知邮件（学生）","论文送审通知邮件（导师）","论文待评阅通知邮件","论文待评阅通知短信","论文审核通知邮件","论文审核未通过通知邮件","论文审核通过通知邮件","评阅意见确认通知邮件","信息导入通知邮件（学生）","信息导入通知邮件（导师）","待办事项通知邮件")
ReDim arrMailId(UBound(arrMailSubject))
Select Case curstep
Case "1"
	ReDim stu_clientstatus(OPRTYPE_COUNT*STUTYPE_COUNT)
	Connect conn
	sql="SELECT * FROM SystemSettings WHERE USE_YEAR="&sem_info(0)&" AND USE_SEMESTER="&sem_info(1)
	GetRecordSetNoLock conn,rs,sql,result
	If result=0 Then	' 本学期无系统设置
		For i=1 To OPRTYPE_COUNT
			stu_startdate(i)=Date
			stu_enddate(i)=Date+7
		Next
		exp_startdate=Date+7
		exp_enddate=Date+14
		tutor_startdate=Date
		tutor_enddate=Date+7
		turn_num=1
		isValid=False
	Else	' 本学期有系统设置
		bSet=True
		For i=1 To OPRTYPE_COUNT
			stu_startdate(i)=rs("STU_"&arrStuOprFlag(i)&"_STARTDATE")
			stu_enddate(i)=rs("STU_"&arrStuOprFlag(i)&"_ENDDATE")
		Next
		exp_startdate=rs("EXP_STARTDATE")
		exp_enddate=rs("EXP_ENDDATE")
		tutor_startdate=rs("TUTOR_STARTDATE")
		tutor_enddate=rs("TUTOR_ENDDATE")
		stu_clientstatus=Split(rs("STU_CLIENT_STATUS"),",")
		turn_num=rs("TURN_NUM")
		For i=1 To UBound(arrMailSubject)
			arrMailId(i)=rs("MAIL_"&i)
		Next
		isValid=rs("VALID")
	End If
	CloseRs rs
	If Not bSet Then
		sql="SELECT TOP 1 * FROM SystemSettings ORDER BY ID DESC"
		GetRecordSetNoLock conn,rs,sql,result
		If result Then ' 显示最新学期的邮件内容
			For i=1 To UBound(arrMailId)
				arrMailId(i)=rs("MAIL_"&i)
			Next
		End If
	End If
	CloseRs rs
	CloseConn conn
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>专业学位论文评阅系统设置</title>
<% useStylesheet(Array("admin", "jeasyui")) %>
<% useScript(Array("jquery", "jeasyui", "common", "systemSettings")) %>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b><%=sem_info(0)%>-<%=sem_info(0)+1%>年度<%=sem_info(2)%>学期专业学位论文评阅系统设置</b><br><%
	If Not bSet Then
%><span style="color:red;font-weight:bold">(请先设置开放时间等属性，方可开放系统)</span><%
	End If %></font>
<form id="fmSettings" method="post" action="?step=2" onsubmit="return chkForm()">
<input type="hidden" name="In_PERIOD_ID" value="<%=sem_info(3)%>">
<table width="900" cellpadding="2" cellspacing="1" bgcolor="dimgray">
<tr bgcolor="ghostwhite">
<td align="left"><%
	Select Case Request("ok")
	Case "0"
%><span style="color:red">更新设置时发生错误，请检查设置参数是否正确</span><br/><%
	Case "1"
%><span style="color:red">设置成功</span><br/><%
	End Select
%></td></tr>
<tr bgcolor="ghostwhite">
<td align="left">其他设置：
	<a href="noticeText.asp">提示文本设置</a>
</td></tr>
<tr bgcolor="ghostwhite">
<td align="center">当前系统状态：<%
	If isValid Then
%>开放<br/><input type="button" name="btnOpenSystem" value="关闭评阅系统" onclick="this.disabled=true;location.href='setSystemStatus.asp?open=1'" /><%
	Else
%>关闭<br/><input type="button" name="btnOpenSystem" value="开放评阅系统" <% If Not bSet Then %>disabled <% Else %>onclick="this.disabled=true;location.href='setSystemStatus.asp?open=0'" <% End If %>/><%
	End If %><input type="button" value="del" onclick="location.href='setSystemStatus.asp?open=101'" style="display:none" /></td></tr>
</td></tr>
<tr bgcolor="ghostwhite">
<td>
	评阅活动：<input id="activity_id" name="activity_id" value="请选择评阅活动…"
	editable="false" style="width: 300px" />
	<a id="btn_add_activity" href="#" class="easyui-linkbutton" data-options="iconCls:'icon-add'">新建评阅活动…</a>
</td>
</tr>
<tr bgcolor="ghostwhite">
<td align="center">
	<table id="activity_period" title="开放时间设置" style="width:90%;height:400px"
		toolbar="#toolbar" idField="id"
		rownumbers="true" fitColumns="true" singleSelect="true">
		<thead>
			<tr>
				<th field="client_type" width="50" disabled>用户类型</th>
				<th field="section" width="50" disabled>环节名称</th>
				<th field="start_time" width="50" editor="{type: 'datetimebox', options: { formatter: Common.formatDateTime, showSeconds: false }}">开始时间</th>
				<th field="end_time" width="50" editor="{type: 'datetimebox', options: { formatter: Common.formatDateTime, showSeconds: false }}">结束时间</th>
				<th field="enabled" width="50" editor="checkbox">是否开放</th>
			</tr>
		</thead>
	</table>
	<div id="toolbar">
		<a href="#" class="easyui-linkbutton" iconCls="icon-add" plain="true" onclick="javascript:$('#dg').edatagrid('addRow')">New</a>
		<a href="#" class="easyui-linkbutton" iconCls="icon-remove" plain="true" onclick="javascript:$('#dg').edatagrid('destroyRow')">Destroy</a>
		<a href="#" class="easyui-linkbutton" iconCls="icon-save" plain="true" onclick="javascript:$('#dg').edatagrid('saveRow')">Save</a>
		<a href="#" class="easyui-linkbutton" iconCls="icon-undo" plain="true" onclick="javascript:$('#dg').edatagrid('cancelRow')">Cancel</a>
	</div>
</td></tr>
<tr bgcolor="ghostwhite"><td colspan=4><p align="center"><input type="submit" value="更改设置"></p></td></tr>
<tr bgcolor="ghostwhite"><td align="left"><font style="font-weight:bold">通知邮件/短信内容设置</font></td></tr>
<tr bgcolor="ghostwhite"><td align="left"><select id="maillist" onchange="switchMailContent(this.selectedIndex)"><option>【请选择】</option><%
	For i=1 To UBound(arrMailSubject)
%><option><%=arrMailSubject(i)%></option><%
	Next
%></select><br />
<span id="mailtip">字段符号:<br/>$stuname - 学生姓名,$stuno - 学号,$stuclass - 学生班级,$stuspec - 所选专业,$stumail - 学生邮箱,<br/>
$subject - 论文题目,$tutorname - 导师姓名,$tutormail - 导师邮箱,$expertname - 专家姓名,<br/>$filename - 审核文件名称/意见类型,$uploadtime - 审核文件上传时间,$evaltext - 导师意见,$postscript - 附注</span></td></tr>
<tr bgcolor="ghostwhite">
<td align="left"><%
	For i=1 To UBound(arrMailSubject)
%><div id="divmailcontent<%=i%>" style="display:none"><% SetEditorWithName "mailcontent"&i,getEmailTemplateContent(arrMailId(i)),170 %></div><%
	Next %>
</td></tr>
<tr bgcolor="ghostwhite"><td align="center"><input type="submit" value="更改设置"></td></tr>
</table></form>
<div id="dialog_add_activity">
	名称：<input id="activity-name" class="easyui-textbox" style="width:200px">
</div>
</center></body>
<script type="text/javascript">
	var Common = {
		curryLoadFilter: function() {
			var args = arguments;
			return function(data) {
				if(data.status !== "ok") {
					$.messager.alert("提示",data.msg,"error");
					return [];
				}
				var ret = data.data;
				for(h in args) {
					ret = typeof args[h] === "function" ? args[h].call(ret) : ret;
				}
				return ret;
			}
		},
		curryOnLoadFailed: function(opr) {
			return function() {
				$.messager.alert("提示", opr+"时出错，请稍后再试。","error");
			}
		},
		formatDateTime: function(date) {
			return date.getFullYear()+"-"+(date.getMonth()+1)+"-"+date.getDay()+" "
				+date.getHours()+":"+date.getMinutes();
		}
	};
	
	$(function() {
		$("#activity_id").combobox({
			url: "../api/get-activities-brief",
			valueField: "id",
			textField: "name",
			loadFilter: Common.curryLoadFilter(Array.prototype.reverse),
			onLoadFailed: Common.curryOnLoadFailed("获取评阅活动列表"),
			onSelect: Common.onComboSelect
		});
		$("#activity_period").datagrid();
		$("#btn_add_activity").bind('click', function() {
			$('#dialog_add_activity').dialog('open');
		});
		$('#dialog_add_activity').dialog({
			title: '新建评阅活动',
			width: 400,
			height: 200,
			closed: true,
			cache: false,
			modal: true
		});

		$('#btnnewturn').click(function() {
			$(this).val('正在执行，请稍候……').attr('disabled',true);
			$(':input[name="turn_num"]').val('<%=turn_num+1%>');
			this.form.submit();
		}).attr('disabled',false);
		$('#btnexport').click(function() {
			$(this).val('正在导出，请稍候……').attr('disabled',true);
			this.form.action='exportReviewStats.asp?fn=<%=sem_info(3)%>_<%=turn_num%>&turn=<%=turn_num%>';
			this.form.submit();
		}).attr('disabled',false);
	});
</script></html><%
Case "2"
	Dim ok
	Dim mail_content,fieldlist
	ReDim mail_content(UBound(arrMailSubject))
	For i=1 To OPRTYPE_COUNT
		stu_startdate(i)=Request.Form("stu_startdate"&i)
		stu_enddate(i)=Request.Form("stu_enddate"&i)
		If Len(stu_startdate(i))=0 Then stu_startdate(i)=Null
		If Len(stu_enddate(i))=0 Then stu_enddate(i)=Null
	Next
	exp_startdate=Request.Form("exp_startdate")
	exp_enddate=Request.Form("exp_enddate")
	If Len(exp_startdate)=0 Then exp_startdate=Null
	If Len(exp_enddate)=0 Then exp_enddate=Null
	tutor_startdate=Request.Form("tutor_startdate")
	tutor_enddate=Request.Form("tutor_enddate")
	If Len(tutor_startdate)=0 Then tutor_startdate=Null
	If Len(tutor_enddate)=0 Then tutor_enddate=Null
	stu_clientstatus="0"
	For i=1 To OPRTYPE_COUNT*STUTYPE_COUNT
		stu_clientstatus=stu_clientstatus&","
		If Request.Form("stu_clientstatus"&i)="on" Then
			stu_clientstatus=stu_clientstatus&"1"
		Else
			stu_clientstatus=stu_clientstatus&"0"
		End If
	Next
	turn_num=Request.Form("turn_num")
	If Not IsNumeric(turn_num) Then
		turn_num=1
	ElseIf turn_num<1 Then
		turn_num=1
	End If
	turn_num=Int(turn_num)
	For i=1 To UBound(arrMailSubject)
		arrMailId(i)=0
		mail_content(i)=Request.Form("mailcontent"&i)
	Next
	arrFieldList=Array("","$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail",_
								"$subject,$expertname,$expertmob,$expertmail,$postscript",_
								"$subject,$expertname,$expertmob,$expertmail,$postscript",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail,$filename,$uploadtime",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail,$filename,$evaltext",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail,$filename,$evaltext",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail,$evaltext",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail,$filename",_
								"$stuname,$stuno,$stuclass,$stuspec,$stumail,$subject,$tutorname,$tutormail,$filename",_
								"$tutorname,$tutormail,$postscript")
	
	Connect conn
	sql="SELECT * FROM SystemSettings WHERE USE_YEAR="&sem_info(0)&" AND USE_SEMESTER="&sem_info(1)
	GetRecordSet conn,rs,sql,result
	On Error Resume Next
	If result=0 Then
		rs.AddNew()
		rs("USE_YEAR")=sem_info(0)
		rs("USE_SEMESTER")=sem_info(1)
	End If
	For i=1 To OPRTYPE_COUNT
		rs("STU_"&arrStuOprFlag(i)&"_STARTDATE")=stu_startdate(i)
		rs("STU_"&arrStuOprFlag(i)&"_ENDDATE")=stu_enddate(i)
	Next
	rs("EXP_STARTDATE")=exp_startdate
	rs("EXP_ENDDATE")=exp_enddate
	rs("TUTOR_STARTDATE")=tutor_startdate
	rs("TUTOR_ENDDATE")=tutor_enddate
	rs("STU_CLIENT_STATUS")=stu_clientstatus
	rs("TURN_NUM")=turn_num
	For i=1 To UBound(arrMailSubject)
		arrMailId(i)=rs("MAIL_"&i)
		If result=0 Or IsNull(arrMailId(i)) Then arrMailId(i)=0
	Next
	rs.Update()
	
	If Err.Number=0 Then ok=1 Else ok=0
	On Error GoTo 0
	
	If ok Then
		sql="UPDATE SystemSettings SET "
		For i=1 To UBound(arrMailSubject)
			template_name=sem_info(0)&"-"&(sem_info(0)+1)&"年度"&sem_info(2)&"学期"&arrMailSubject(i)
			arrMailId(i)=updateEmailTemplate(arrMailId(i),template_name,arrMailSubject(i),mail_content(i),arrFieldList(i))
			If i>1 Then sql=sql&","
			sql=sql&"MAIL_"&i&"="&arrMailId(i)
		Next
		sql=sql&" WHERE USE_YEAR="&sem_info(0)&" AND USE_SEMESTER="&sem_info(1)
		conn.Execute sql
	End If
	
	CloseRs rs
	CloseConn conn
%><body><form id="ret" method="post" action="?step=1"><input type="hidden" name="ok" value="<%=ok%>" />
<input type="hidden" name="In_PERIOD_ID" value="<%=sem_info(3)%>">
</form>
<script type="text/javascript">
document.all.ret.submit();
</script></body><%
End Select
%>
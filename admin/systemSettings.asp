<!--#include file="../inc/global.inc"-->
<!--#include file="../inc/setEditor.asp"-->
<!--#include file="../inc/ckeditor/ckeditor.asp"-->
<!--#include file="../inc/ckfinder/ckfinder.asp"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")

Dim configs:Set configs=getSystemConfigs()
Dim status:status=configs("Status")
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>专业学位论文评阅系统设置</title>
<% useStylesheet "admin", "global", "jeasyui" %>
<% useScript "jquery", "jeasyui", "common", "systemSettings" %>
</head>
<body bgcolor="ghostwhite" style="overflow-y: scroll">
<center><font size=4><b>专业学位论文评阅系统基本设置</b></font>
<table width="900" cellpadding="2" cellspacing="1" bgcolor="dimgray">
<tr bgcolor="ghostwhite"><td align="left">其他设置：
	<a href="noticeText.asp">提示文本设置</a>
</td></tr>
<tr bgcolor="ghostwhite">
<td>
	<div style="margin: 5px 8px">
		<div id="switch_system_status" style="width:100px;"></div>
	</div>
	<div style="margin: 5px auto">
		<input class="easyui-combobox" id="activity_id" name="activity_id" />
		<a class="easyui-linkbutton" id="btn_edit_activity" href="#" data-options="disabled:true,iconCls:'icon-edit'" title="编辑评阅活动"></a>
		<a class="easyui-linkbutton" id="btn_remove_activity" href="#" data-options="disabled:true,iconCls:'icon-cancel'" title="删除评阅活动"></a>
		<a class="easyui-linkbutton" id="btn_add_activity" href="#" data-options="iconCls:'icon-add'" title="新增评阅活动"></a>
	</div>
	<div id="settings" style="display: none">
		<div style="height: 5px"></div>
		<table class="easyui-treegrid" id="treegrid_activity_period">
			<thead>
				<tr>
					<th field="name" width="50" disabled>环节名称</th>
					<th field="start_time" width="50" editor="datetimebox">开始时间</th>
					<th field="end_time" width="50" editor="datetimebox">结束时间</th>
					<th field="_enabled" width="50" editor="checkbox">是否开放</th>
				</tr>
			</thead>
		</table>
		<div style="height: 5px"></div>
		<div id="panel_mail_templates" style="padding: 5px">
			<div><p>
	字段符号:<br/>$stuname - 学生姓名,$stuno - 学号,$stuclass - 学生班级,$stuspec - 所选专业,$stumail - 学生邮箱,<br/>
	$subject - 论文题目,$tutorname - 导师姓名,$tutormail - 导师邮箱,$expertname - 专家姓名,<br/>$filename - 审核文件名称/意见类型,$uploadtime - 审核文件上传时间,$evaltext - 导师意见,$postscript - 附注
			</p></div>
			<div style="display: flex; flex-direction: row">
				<div style="flex-basis: 390px; height: 425px; overflow-y: scroll">
					<ul id="tree_mail_template_list"></ul>
				</div>
				<div>
					<form id="form_mail_template" method="post">
						<% SetEditorWithName "mail_template_content","",160 %>
					</form>
				</div>
			</div>
		</div>
	</div>
</td></tr></table>
<div id="dialog_activity">
	<form id="form_activity" method="post">
		<div>
			<input class="easyui-combobox" id="activity_semester" name="semester" />
		</div>
		<div>
			<input class="easyui-combobox" id="activity_stu_type" name="stu_type" />
		</div>
		<div>
			<input class="easyui-textbox" id="activity_name" name="name" />
		</div>
		<div>
			<div id="activity_switch_is_open" style="width:100px;"></div>
		</div>
		<div style="text-align: center;">
			<a href="#" class="easyui-linkbutton" id="btn_submit_dialog_activity" iconCls="icon-save">确 定</a>
			<a href="#" class="easyui-linkbutton" id="btn_close_dialog_activity" iconCls="icon-cancel">关 闭</a>
		</div>
	</form>
</div></center>
<script type="text/javascript">
	$(document).data({
		system_status: "<%=status%>"
	});
</script>
</body></html>
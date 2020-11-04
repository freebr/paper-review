function addStatusIndicator() {
    for(var i=0;i<this.length;++i) {
        this[i].name=({ false: '○', true: '●' })[this[i].is_open]+this[i].name;
    }
    return this;
}

function onSwitchSystemStatusChanged(checked) {
    $.post("../api/set-system-status",
        {
            status: checked ? 'open' : 'closed'
        },
        function (data) {
            if (data.status==='ok') {
			    $.messager.show({ msg: "操作成功", title: "设置系统状态", timeout: 2000 });
            }
        }
    );
}

function onActivityIdSelect(item) {
    if ($("#treegrid_activity_period").data("activity_id")===item.id) return;
    $("#btn_edit_activity, #btn_remove_activity").linkbutton("enable");
    $("#treegrid_activity_period").treegrid("load", {
        activity_id: item.id
    }).data("activity_id", item.id).treegrid("getPanel").panel("open");
    $("#treegrid_activity_period").resize();
	$("#tree_mail_template_list").tree({
		url: "../api/get-mail-template-list?activity_id="+item.id
	});
    $("#panel_mail_templates").panel("open");
    $("#settings").show();
}

function onActivityPeriodsBeforeLoad(row, param) {
    return !!param.activity_id;
}

function onActivityPeriodsLoadSuccess(row, data) {
    if (data.length>1) {
        $(this).treegrid("collapseAll");
    }
}

function activityPeriodDataHandler(is_update) {
    for(var i=0;i<this.length;++i) {
        var children=this[i].children;
        if (children===undefined) continue;
        for(var j=0;j<children.length;++j) {
            var row=children[j];
            row.start_time=row.start_time.replace(/\//g, "-");
            row.end_time=row.end_time.replace(/\//g, "-");
            if (is_update) row.enabled=!!row._enabled;
            row._enabled=$("<input type='checkbox'>").attr("checked", row.enabled)[0].outerHTML;
        }
    }
    return this;
}

function onActivityPeriodsClickCell(field, row) {
    var fn=$(this).treegrid.bind($(this));
    fn("getSelected") && fn("endEdit", fn("getSelected").id);
    if (field==="name" || row.children) return;
    fn("beginEdit", row.id);
    $(fn("getEditor", { id:row.id, field:"_enabled" }).target)
        .prop({ checked: row.enabled, value: true });
}

function onActivityPeriodsAfterEdit(row, changes) {
    const reZero=/(?=\s)0+:0+(:0+)?\b/;
    row.start_time=row.start_time.replace(reZero, '');
    row.end_time=row.end_time.replace(reZero, '');
    if (row.end_time.match(/^\d{4}-\d{1,2}-\d{1,2}$/)) row.end_time += ' 23:59:59';
    data=activityPeriodDataHandler.call(Array({ children: [row]}), true);
    $(this).treegrid("update", { id: row.id, row: data[0].children[0] });
    $.post("../api/set-activity-section-periods", {
        activity_id: $("#activity_id").val(),
        stu_type: row.stu_type,
        client_type: row.client_type,
        section_id: row.section_id,
        start_time: row.start_time,
        end_time: row.end_time,
        enabled: row.enabled
    });
}

function onAddActivitySemesterChange() {
    if ($("#dialog_activity").data("type")==="add") {
        return generateActivityName();
    }
}
function onAddActivityStuTypeChange() {
    if ($("#dialog_activity").data("type")==="add") {
        return generateActivityName();
    }
}
function generateActivityName() {
    var semester_name=$("#activity_semester").combobox("getText");
    var stu_types=$("#activity_stu_type").combobox("getValues");
    var data=$("#activity_stu_type").combobox("getData");
    var stu_type_name;
    if (stu_types.indexOf('5')>-1&&stu_types.indexOf('6')>-1&&
        stu_types.indexOf('7')>-1&&stu_types.indexOf('9')>-1) {
        stu_type_name='3M、EMBA';
    } else if (stu_types.length===3&&stu_types.indexOf('5')>-1&&
        stu_types.indexOf('6')>-1&&stu_types.indexOf('9')>-1) {
        stu_type_name='3M';
    } else {
        stu_type_name=stu_types.map(function(item) {
            return data.filter(function(r) {
                return r.id == item;
            })[0].name;
        }).toString();
    }
    $("#activity_name").textbox("setValue", semester_name+stu_type_name+"论文评阅");
}

function onMailTemplateListClick(node) {
	if (node.mail_content===undefined) return;
	var instance=CKEDITOR.instances.mail_template_content;
	var fn=function() {
		instance.template=node;
		instance.setData(node.mail_content, {
			callback: function() {
				this.resetDirty();
			}
		});
	}
	if (!instance.checkDirty()) {
		fn();
	} else {
		$.messager.confirm("提示", "模板已修改，是否保存？",
			function(r) {
				r ? onMailTemplateSubmit(fn) : fn();
			}
		);
	}
}
function onMailTemplateSubmit(fn) {
	var instance=CKEDITOR.instances.mail_template_content;
	var new_mail_content=instance.getData();

    $.post("../api/set-mail-template",
        {
            id: instance.template.id,
            mail_content: new_mail_content
        },
        function(data) {
            if (data.status==='ok') {
                $.messager.show({ msg: "保存成功", title: "保存通知邮件（短信）模板", timeout: 2000 });
                instance.template.mail_content=new_mail_content;
                instance.resetDirty();
                fn&&fn();
            } else {
                $.messager.show({ msg: "保存失败："+data.msg, title: "保存通知邮件（短信）模板", timeout: 2000 });
            }
        }
    );
}

$(function() {
    $(this).data({
        "clipboard": null
    });
    $("#switch_system_status").switchbutton({
		label: "系统状态：",
        labelWidth: 90,
        onText: "开放",
        offText: "关闭",
        checked: $(this).data("system_status") === 'open',
        onChange: onSwitchSystemStatusChanged
    });
    $("#activity_id").combobox({
        url: "../api/get-activities",
        valueField: "id",
        textField: "name",
        label: "评阅活动：",
        labelWidth: 90,
        labelAlign: "right",
        width: 400,
        value: "请选择评阅活动…",
        loadFilter: Common.curryLoadFilter(Array.prototype.reverse, addStatusIndicator),
        onLoadFailed: Common.curryOnLoadFailed("获取评阅活动列表"),
        onSelect: onActivityIdSelect
    });
    $("#treegrid_activity_period").treegrid({
		url: "../api/get-activity-periods",
		width: "100%",
		height: 450,
		title: "开放时间设置",
		iconCls: "icon-schedule",
        animation: true,
        idField: "id",
        treeField: "name",
        fitColumns: true,
		singleSelect: true,
        collapsible: true,
        striped: true,
        closed: true,
        toolbar: [{
            iconCls: "icon-save",
            plain: true,
            text: "保存修改",
            handler: function() {
                $activity_period=$("#treegrid_activity_period");
                var sel=$activity_period.treegrid("getSelected");
                if (!sel||!sel.client_type) return;
                $activity_period.treegrid("endEdit",sel.id);
                $.messager.show({ msg: "保存成功", title: "保存修改", timeout: 2000 });
            }
        }, {
            iconCls: "icon-cancel",
            plain: true,
            text: "放弃修改",
            handler: function() {
                $activity_period=$("#treegrid_activity_period");
                var sel=$activity_period.treegrid("getSelected");
                if (!sel||!sel.client_type) return;
                $.messager.confirm("提示", "确定要放弃所做的修改吗？",
                    function(r) {
                        if (!r) return;
                        $activity_period.treegrid("cancelEdit",sel.id);
                    }
                );
            }
        }, "-", {
            iconCls: "icon-copy",
            plain: true,
            text: "拷贝开放时间",
            handler: function() {
                var sel=$("#treegrid_activity_period").treegrid("getSelected");
                if (!sel) {
                    $.messager.show({ msg: "请选择要拷贝的记录。", title: "拷贝开放时间" });
                    return;
                }
                $(document).data("clipboard", { type: sel.children ? 0 : 1, row: sel });
                $.messager.show({ msg: "记录已拷贝，请选择要粘贴的记录，点击“粘贴开放时间”完成粘贴。",
                    title: "拷贝开放时间" });
            }
        }, {
            iconCls: "icon-paste",
            plain: true,
            text: "粘贴开放时间",
            handler: function() {
                var clipboard=$(document).data("clipboard");
                if (!clipboard) return;
                var $treegrid=$("#treegrid_activity_period");
                var sel=$treegrid.treegrid("getSelected");
                var type=sel.children ? 0 : 1;
                if (sel.id===clipboard.id) return;
                if (type!==clipboard.type) {
                    $.messager.show({ msg: "所选记录与剪贴板中的记录类型不同，无法粘贴。",
                        title: "粘贴开放时间" });
                    return;
                }
                if (type===0) {	// 学生类型
                    for (var i=0;i<sel.children.length;++i) {
                        $treegrid.treegrid("update", {
                            id: sel.children[i].id,
                            row: {
                                start_time: clipboard.row.children[i].start_time,
                                end_time: clipboard.row.children[i].end_time,
                                enabled: clipboard.row.children[i].enabled
                            }
                        });
                        onActivityPeriodsClickCell.call($treegrid,"",sel.children[i]);
                        $treegrid.treegrid("endEdit", sel.children[i].id);
                    }
                } else {	// 环节
                    $treegrid.treegrid("update", {
                        id: sel.id,
                        row: {
                            start_time: clipboard.row.start_time,
                            end_time: clipboard.row.end_time,
                            enabled: clipboard.row.enabled
                        }
                    });
                    onActivityPeriodsClickCell.call($treegrid,"",sel);
                    $treegrid.treegrid("endEdit", sel.id);
				}
				$.messager.show({ msg: "粘贴成功", title: "粘贴开放时间" });
            }
        }],
        loadMsg: "正在加载……",
        loadFilter: Common.curryLoadFilter(activityPeriodDataHandler),
        onBeforeLoad: onActivityPeriodsBeforeLoad,
        onLoadSuccess: onActivityPeriodsLoadSuccess,
        onClickCell: onActivityPeriodsClickCell,
        onAfterEdit: onActivityPeriodsAfterEdit
    });
    $("#btn_edit_activity").bind("click", function() {
        if ($(this).linkbutton("options").disabled) return;
        $("#dialog_activity").dialog({
            title: "编辑评阅活动"
        }).dialog("open").data("type", "edit");
        $.post("../api/get-activity",
            {
                id: $("#activity_id").val()
            },
            function(data) {
                if (data.status!=='ok') {
                    $.messager.alert("提示", data.msg, "error");
                    return;
                }
                var item=data.data;
                $("#activity_stu_type").combobox({ disabled: true });
                $("#form_activity").form("clear").form("load", {
                    name: item.name,
                    semester: item.semester_id,
                    stu_type: item.stu_type_id,
                });
                $("#activity_switch_is_open").switchbutton(item.is_open ? 'check' : 'uncheck');
            }
        );
    });
    $("#btn_add_activity").bind("click", function() {
        $("#dialog_activity").dialog({
            title: "新增评阅活动"
        }).dialog("open").data("type", "add");
        $("#activity_stu_type").combobox({ disabled: false });
        $("#form_activity").form("clear").form("load", {
            semester: $("#activity_semester").combobox("getData")[0].period_id,
        });
        $("#activity_switch_is_open").switchbutton("check");
    });
    $("#btn_remove_activity").bind("click", function() {
        var id=$("#activity_id").val();
        var name=$("#activity_id").combobox("getText");
        if (!id) return;
        var fn=function(r) {
            if (!r) return;
            $.messager.progress();
            $.post("../api/remove-activity",
                {
                    id: id
                },
                function(data) {
                    if (data.status==="ok") {
                        $.messager.show({ msg: "操作成功", title: "删除评阅活动", timeout: 2000 });
                        $("#activity_id").combobox("reload").combobox("reset");
                        $("#treegrid_activity_period").treegrid("loadData", { status: "ok", data: [] });
                    } else {
                        $.messager.show({ msg: "操作失败："+data.msg, title: "删除评阅活动", timeout: 2000 });
                    }
                    $.messager.progress("close");
                }
            );
        }
        $.messager.confirm("提示", "确定要删除评阅活动【"+name+"】吗？", fn);
    });
    $("#btn_submit_dialog_activity").bind("click", function() {
        var id=$("#activity_id").val();
        var title=$("#dialog_activity").dialog("options").title;
        $.messager.progress();
        $("#form_activity").form("submit", {
            url: ({ add: "../api/add-activity",
                    edit: "../api/edit-activity" })
                 [$("#dialog_activity").data("type")],
            onSubmit: function(param) {
                var isValid = $(this).form("validate");
                if (!isValid) {
                    $.messager.show({ msg: "您所提交的信息有误，请检查。", title: title, timeout: 2000 });
                    $.messager.progress("close");
                }
                param.id=id;
                param.is_open=$("#activity_switch_is_open").switchbutton("options").checked;
                return isValid;
            },
            success: function(data) {
                data=JSON.parse(data);
                if (data.status==="ok") {
                    $.messager.show({ msg: "操作成功", title: title, timeout: 2000 });
                    $("#activity_id").combobox("reload").combobox("setValue", id);
                } else {
                    $.messager.show({ msg: "操作失败："+data.msg, title: title, timeout: 2000 });
                }
                $("#dialog_activity").dialog("close");
                $.messager.progress("close");
            }
        });
    });
    $("#btn_close_dialog_activity").bind("click", function() {
        $("#dialog_activity").dialog("close");
    });
    $("#dialog_activity").dialog({
        title: "新建评阅活动",
        width: 500,
        top: 100,
        closed: true,
        cache: false,
        modal: true
    });

    $("#activity_semester").combobox({
        url: "../api/get-semester-list",
        label: "适用学期",
        labelAlign: "right",
        missingMessage: "请选择评阅活动适用学期",
        required: true,
        textField: "period_name",
        valueField: "period_id",
        width: 450,
        loadFilter: Common.curryLoadFilter(),
        onChange: onAddActivitySemesterChange
    });
    $("#activity_stu_type").combobox({
        url: "../api/get-stu-type-list",
        label: "适用学生类型",
        labelAlign: "right",
        missingMessage: "请选择评阅活动适用学生类型",
        multiple: true,
        required: true,
        textField: "name",
        valueField: "id",
        value: null,
        width: 450,
        panelHeight: 200,
        loadFilter: Common.curryLoadFilter(),
        onChange: onAddActivityStuTypeChange
    });
    $("#activity_name").textbox({
        label: "评阅活动名称",
        labelAlign: "right",
        required: true,
        missingMessage: "请填写评阅活动名称",
        width: 450
    });
    $("#activity_switch_is_open").switchbutton({
		label: "状态",
        labelAlign: "right",
        onText: "开放",
        offText: "关闭",
        width: 80
    });

	$("#panel_mail_templates").panel({
		width: "100%",
		height: 560,
		title: "通知邮件（短信）模板设置",
		iconCls: "icon-email",
		closed: true,
		collapsible: true
	});
	$("#tree_mail_template_list").tree({
		loadFilter: Common.curryLoadFilter(),
		onClick: onMailTemplateListClick
	});
	$("#btn-save-template").linkbutton({
		iconCls: 'icon-save',
		plain: true
	});
	$("#btn-cancel-template").linkbutton({
		iconCls: 'icon-cancel',
		plain: true
	});
	CKEDITOR.on('instanceReady', function(ev) {
		ev.editor.on('beforeCommandExec', function(event){
			if (event.data.name==='save') {
				onMailTemplateSubmit();
				return false;
			}
		});
	});
});
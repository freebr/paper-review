function submitForm(fm,action,enctype,data) {
	if(typeof(fm.size)==='function')fm=fm[0];
	if(!!action)fm.action=action;
	if(fm.id=='query_nocheck') {
		var ctls=['TEACHTYPE_ID','ENTER_YEAR','CLASS_ID'];
		for(i=0;i<ctls.length;i++) {
			var selindex=fm[ctls[i]].selectedIndex;
			if(selindex<=0) {
				fm['In_'+ctls[i]].value='0';
			} else {
				fm['In_'+ctls[i]].value=fm[ctls[i]].options[selindex].value;
			}
		}
	}
	fm.encoding=(!enctype)?"application/x-www-form-urlencoded":enctype;
	if(data instanceof Object) {
		for(var key in data) {
			$(fm).append($("<input type='hidden'>").attr({ name: key, value: data[key] }));
		}
	}
	fm.submit();
	return false;
}
function matchReviewer(fm,tid) {
	submitForm(fm,"matchReviewer.asp?tid="+tid);
}
function notifyReviewer(fm,tid) {
	submitForm(fm,"notifyReviewer.asp?tid="+tid);
}
function batchFetchDocument(fm) {
	var ids='';
	$(fm).find(':checked[name="sel"]').each(function(index,item){
		ids+=(ids.length?',':'')+item.value;
	});
	!window.tabmgr?window.open('batchFetchDocument.asp?sel='+ids):
		window.tabmgr.newTab('/PaperReview/admin/batchFetchDocument.asp?sel='+ids);
}
function batchUpdatePaper(fm) {
	$(fm).find(':hidden[name="review_display_status"]')
			 .val($('select[name="selreviewfilestat"]').val());
	submitForm(fm,"batchUpdatePaper.asp");
}
function showAllRecords(fm) {
	submitForm(fm,"paperList.asp?showAll");
}
function showPaperDetail(id,usertype) {
	var client=['admin','student','tutor','expert'];
	!window.tabmgr?window.open('paperDetail.asp?tid='+id,'thesis'+id):
		window.tabmgr.newTab('/PaperReview/'+client[usertype]+'/paperDetail.asp?tid='+id);
	return false;
}
function showStudentProfile(id,usertype) {
	var client=['admin','student','tutor','expert'];
	!window.tabmgr?window.open('studentProfile.asp?id='+id):
		window.tabmgr.goTo('/PaperReview/'+client[usertype]+'/studentProfile.asp?id='+id,"学生基本资料",true);
	return false;
}
function showTeacherProfile(id) {
	var url="http://oa.cnsba.com/teacher_resume.asp?id="+id;
	window.open(url,'teacher'+id);
	return false;
}
function showExpertProfile(id) {
	!window.tabmgr?window.open('expertProfile.asp?id='+id,'expert'+id):
		window.tabmgr.newTab('/PaperReview/admin/expertProfile.asp?id='+id);
	return false;
}
function rollback(tid,user,opr) {
	if(user!=0&&user!=1&&user!=2&&user!=3) return false;
	var msg_templ=["确实要撤销这个文件吗？","确实要撤销这名专家的评阅操作吗？",
					"确实要撤销导师的审核操作吗？","确实要撤销该项操作吗？"];
	var msg_templ_ps=[["开题报告表","开题论文","中期考核表","中期论文","预答辩意见书","预答辩论文","最新上传的送检论文","送审论文","答辩论文","教指委盲评论文","定稿论文","答辩审批材料"],
							["第一位专家的评阅书和评阅意见","第二位专家的评阅书和评阅意见"],
							["导师对表格材料的审核意见","导师的送检意见","导师的送审意见","导师对答辩论文的意见"],
							["该学生的所有送检论文和送检报告","该论文的评阅专家匹配信息及评阅结果","该论文的答辩安排信息","该论文的答辩委员会修改意见","该论文的教指委委员匹配结果","第一位教指委委员的修改意见","第二位教指委委员的修改意见","该论文的学院学位评定分会修改意见"]]
	var msg=msg_templ[user]+msg_templ_ps[user][opr]+"将会被删除且不可恢复！"
	if (confirm(msg)) {
		submitForm(document.all.fmDetail,"rollback.asp",null,{ tid: tid, user: user, rollback_opr: opr });
		return true;
	}
	return false;
}
function deleteDetectResult(tid,id,delete_type) {
	var msg=["确实要删除该检测报告吗（论文将保留）？","确实要删除该条检测记录吗﹙论文和报告将被删除﹚？"];
	if (confirm(msg[delete_type])) {
		submitForm(document.all.fmDetail,"delDetectResult.asp",null,{ tid: tid, id: id, delete_type: delete_type });
		return true;
	}
	return false;
}
function modifyReview(tid,rid) {
	submitForm(document.all.fmDetail,"extra/paperDetail.asp",null,{ tid: tid, rev: rid });
	return false;
}
function checkLength(txt,len) {
	var tip=$('#'+txt.name+'_tip');
	if (txt.value.length>len) {
		tip.html('<font color="red">已超出&nbsp;'+(txt.value.length-len)+'&nbsp;字</font>');
	} else {
		tip.html('<font color="blue">还可填写&nbsp;'+(len-txt.value.length)+'&nbsp;字</font>');
	}
}
function getFileTypeByAuditType(audit_type) {
	// value 与 global.inc 中的 arrDefaultFileListName() 对应
	return ({
		'1': 1,
		'2': 3,
		'3': 5,
		'4': 7,
		'5': 8,
		'6': 15,
		'7': 0,
		'8': 10,
		'9': 0,
		'10': 11
	})[audit_type] || 0;
}
function initAuditRecordsDataGrid($el, paper_id) {
	$el.datagrid({
		url: "../api/get-audit-records?id="+paper_id,
		columns: [[
			{ field: 'id', title: '#', width: 50, align: 'center',
				formatter: function(value, row, index) {
					var num = $el.datagrid("options").sortOrder == 'asc' ? index + 1 :
						$el.datagrid("getRows").length - index;
					return "<p align='center'>" + num + "</p>";
				}
			},
			{ field: 'audit_time', title: '审核时间', width: 150, align: 'center', sortable: true,
				sorter: function(a, b) {
					return new Date(a) - new Date(b);
				}
			},
			{ field: 'audit_type_name', title: '审核类型', width: 300, align: 'center' },
			{ field: 'audit_file', title: '审核文件', width: 150, align: 'center',
				formatter: function(value, row, index) {
					if (!row.audit_file) return;
					var type = getFileTypeByAuditType(row.audit_type);
					return '<a class="resc" href="fetchDocument.asp?store=audit&type='+type+'&id='+row.id+'" target="_blank">点击下载</a>';
				}
			},
			{ field: 'auditor_name', title: '审核人', width: 190, align: 'center' },
			{ field: 'is_passed', title: '审核结果', width: 125, align: 'center',
				formatter: function(value, row, index) {
					var class_name = value ? "accepted-sign" : "rejected-sign";
					var sign = value ? "√" : "×";
					return "<span class=" + class_name + ">" + sign + "</span>";
				}
			},
		]],
		view: $.extend({}, $.fn.datagrid.defaults.view, {
			renderRow: function(target, fields, frozen, rowIndex, rowData) {
				if (rowData.id == 0) {
					return "<p>暂无审核记录</p>";
				}
				var code = [
					$.fn.datagrid.defaults.view.renderRow.apply(this, arguments)
				];
				code.push(
					"</tr><tr><td colspan='"+fields.length+"' class='record-comment'><p>",
					rowData.comment,
					"</p></td>"
				);
				return code.join("");
			}
		}),
		height: 300,
		remoteSort: false,
		singleSelect: true,
		sortName: 'audit_time',
		sortOrder: 'desc',
		loadFilter: Common.curryLoadFilter()
	});
}
function initReviewRecordsDataGrid($el, paper_id, admin) {
	admin = admin || false;
	var display_status_values = [
		{ value: 0, text: "不开放显示" },
		{ value: 1, text: "仅向导师显示" },
		{ value: 2, text: "完全开放显示" }
	];
	var columns = [[
		{ field: 'review_order_text', title: '评阅顺序', width: 80, align: 'center' },
		{ field: 'review_time', title: '评阅时间', width: 100, align: 'center', sortable: true,
			sorter: function(a, b) {
				return new Date(a) - new Date(b);
			}
		},
		{ field: 'overall_rating_text', title: '总体评价', width: 60, align: 'center' },
		{ field: 'defence_opinion_text', title: '评审结果', width: admin ? 100 : 250, align: 'center' },
		{ field: 'review_file', title: '评阅书', width: 130, align: 'center',
			formatter: function(value, row, index) {
				return '<a class="resc" href="fetchDocument.asp?store=review&type=18&id='+row.id+'" target="_blank">点击下载</a>';
			}
		},
		{ field: 'display_status', title: '是否显示', width: 140, align: 'center',
			formatter: function(value, row, index) {
				value=parseInt(value);
				var ret=display_status_values.filter(function(item) {
					return value===item.value;
				});
				return ret.length ? ret[0].text : admin ? "【请选择】" : "未设置";
			},
			editor: !admin ? null : {
				type: 'combobox',
				options: {
					data: display_status_values,
					panelHeight: 140,
					onChange: onReviewRecordsDisplayStatusChange.bind($el),
				}
			}
		},
		{ field: 'display_status_modified_by_name', title: '显示状态修改人', width: 140, align: 'center' }
	]];
	if (admin) {
		columns[0].splice(1,0,{ field: 'reviewer_name', title: '评阅人', width: 100, align: 'center',
			formatter: function(value, row, index) {
				var h='<a href="#" onclick="return showExpertProfile('+row.reviewer_id+')">'+value+'</a>';
				if (row.creator!==row.reviewer_id) {
					h+='<br />由['+row.creator_name+']操作';
				}
				return h;
			}
		});
		columns[0].push(
			{
				field: 'operation', title: '操作', width: 50, align: 'center',
				formatter: function(value, row, index) {
					return "<a href='#' onclick='return onReviewRecordDelete($(\"#"+$el.attr('id')+"\"),\""+row.id+"\");'>删除</a>";
				}
			}
		);
	}
	$el.datagrid({
		url: "../api/get-review-records?id="+paper_id,
		columns: columns,
		view: $.extend({}, $.fn.datagrid.defaults.view, {
			renderRow: function(target, fields, frozen, rowIndex, rowData) {
				if (rowData.id == 0) {
					return "<p>暂无送审记录</p>";
				}
				var code = [
					$.fn.datagrid.defaults.view.renderRow.apply(this, arguments)
				];
				if ($el.data("load")) {
					code.push(
						"</tr><tr><td colspan='"+fields.length+"' class='record-comment'><p>",
						rowData.comment,
						"</p></td>"
					);
				}
				return code.join("");
			},
			onAfterRender: function(target) {
				$el.data("load", false);
			}
		}),
		height: 200,
		remoteSort: false,
		singleSelect: true,
		sortName: 'review_time',
		sortOrder: 'asc',
		loadFilter: Common.curryLoadFilter(),
		onClickCell: onReviewRecordsSelect
	}).data("load", true);
}
function onReviewRecordsSelect(index, data) {
	$(this).datagrid("endEdit", $(this).data("selected"))
		   .datagrid("beginEdit", index);
	$(this).data("selected", index);
}
function onReviewRecordsDisplayStatusChange(newValue, oldValue) {
	if (!oldValue.length) return;
	var $el = $(this);
	var row = $el.datagrid("getSelected");
	if (!row) return;
	$.post("../api/set-review-record-display-status",
		{
			id: row.id,
			display_status: parseInt(newValue)
		},
		function(data) {
			if (data.status==='ok') {
				$el.data("load", true);
				$el.datagrid("reload");
				$.messager.show({ msg: "保存成功", title: "设置评阅记录显示状态", timeout: 2000 });
			} else {
				$.messager.show({ msg: "保存失败："+data.msg, title: "设置评阅记录显示状态", timeout: 2000 });
			}
		}
	);
}
function onReviewRecordDelete($el, id) {
	if (!confirm("确定删除该条评阅记录吗？将不可恢复！")) return false;
	$.post("../api/delete-review-record",
		{
			id: id
		},
		function(data) {
			if (data.status==='ok') {
				$el.data("load", true);
				$el.datagrid("reload");
				$.messager.show({ msg: "删除成功", title: "删除评阅记录", timeout: 2000 });
			} else {
				$.messager.show({ msg: "删除失败："+data.msg, title: "删除评阅记录", timeout: 2000 });
			}
		}
	);
	return false;
}
function initDefencePlanDataGrid($el, defence_time, defence_place, defence_members, defence_memo) {
	defence_members = defence_members.split('|');
	$el.datagrid({
		title: "答辩安排",
		iconCls: "icon-schedule",
		columns: [[
			{ field: 'defence_time', title: '答辩时间', width: 160, align: 'center' },
			{ field: 'defence_place', title: '答辩地点', width: 120, align: 'center' },
			{ field: 'defence_chairman', title: '答辩主席', width: 80, align: 'center' },
			{ field: 'defence_members', title: '答辩委员', width: 130, align: 'center' },
			{ field: 'defence_secretary', title: '答辩秘书', width: 80, align: 'center' },
			{ field: 'defence_memo', title: '答辩委员工作单位', width: 500, align: 'center' }
		]],
		data: [{
			defence_time: defence_time,
			defence_place: defence_place,
			defence_chairman: defence_members[0],
			defence_members: defence_members[1],
			defence_secretary: defence_members[2],
			defence_memo: defence_memo
		}],
		height: 150,
		autoRowHeight: true
	}).datagrid("autoSizeColumn", "defence_memo");
}
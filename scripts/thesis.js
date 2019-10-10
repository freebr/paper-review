function submitForm(fm,action,enctype,data) {
	if(typeof(fm.size)==='function')fm=fm[0];
	if(!!action)fm.action=action;
	if(fm.name=='query_nocheck') {
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
function chooseExpert(fm,tid) {
	submitForm(fm,"chooseExpert.asp?tid="+tid);
}
function notifyExpert(fm,tid) {
	submitForm(fm,"notifyExpert.asp?tid="+tid);
}
function batchFetchDocument(fm) {
	var ids='';
	$(fm).find(':checked[name="sel"]').each(function(index,item){
		ids+=(ids.length?',':'')+item.value;
	});
	tabmgr.goTo('/ThesisReview/admin/batchFetchDocument.asp?sel='+ids,'批量下载表格/论文',true);
}
function batchUpdateThesis(fm) {
	$(fm).find(':hidden[name="review_display_status"]')
			 .val($('select[name="selreviewfilestat"]').val());
	submitForm(fm,"batchUpdateThesis.asp");
}
function showAllRecords(fm) {
	submitForm(fm,"thesisList.asp?showAll");
}
function showThesisDetail(id,usertype) {
	var client=['admin','student','tutor','expert'];
	!window.tabmgr?window.open('thesisDetail.asp?tid='+id,'thesis'+id):
		window.tabmgr.newTab('/ThesisReview/'+client[usertype]+'/thesisDetail.asp?tid='+id);
	return false;
}
function showExpertProfile(id) {
	!window.tabmgr?window.open('expertProfile.asp?id='+id,'expert'+id):
		window.tabmgr.newTab('/ThesisReview/admin/expertProfile.asp?id='+id);
	return false;
}
function rollback(tid,user,opr) {
	if(user!=0&&user!=1&&user!=2&&user!=3) return false;
	var msg_templ=["确实要撤销这个文件吗？","确实要撤销这名专家的评阅操作吗？",
					"确实要撤销导师的审核操作吗？","确实要撤销该项操作吗？"];
	var msg_templ_ps=[["开题报告表","开题论文","中期检查表","中期论文","预答辩申请表","预答辩论文","最新上传的送检论文","送审论文","答辩论文","定稿论文","答辩审批材料"],
							["第一位专家的评阅书和评阅意见","第二位专家的评阅书和评阅意见"],
							["导师对表格材料的审核意见","导师的送检意见","导师的送审意见","导师对答辩论文的意见"],
							["该学生的所有送检论文和送检报告","该论文的匹配专家结果","该论文的答辩安排信息","该论文的答辩委员会修改意见","该论文的教指会分会修改意见"]]
	var msg=msg_templ[user]+msg_templ_ps[user][opr]+"将会被删除且不可恢复！"
	if (confirm(msg)) {
		submitForm(document.all.fmDetail,"rollback.asp",null,{ tid: tid, user: user, rollback_opr: opr });
		return true;
	}
	return false;
}
function deleteDetectResult(tid,hash,delete_type) {
	var msg=["确实要删除该送检报告吗（送检论文仍将保留）？","确实要删除该条送检记录吗﹙送检论文和报告将被删除﹚？"];
	if (confirm(msg[delete_type])) {
		submitForm(document.all.fmDetail,"delDetectResult.asp",null,{ tid: tid, hash: hash, delete_type: delete_type });
		return true;
	}
	return false;
}
function modifyReview(tid,rid) {
	submitForm(document.all.fmDetail,"extra/thesisDetail.asp",null,{ tid: tid, rev: rid });
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
function submitForm(fm,action,enctype) {
	if(typeof(fm.size)==='function')fm=fm[0];
	if(!!action)fm.action=action;
	if(fm.name=='query_nocheck') {
		var ctls=['TEACHTYPE_ID','CLASS_ID','ENTER_YEAR'];
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
	fm.submit();
	return false;
}
function chooseExpert(fm,tid) {
	submitForm(fm,"chooseExpert.asp?tid="+tid);
}
function notifyExpert(fm,tid) {
	submitForm(fm,"notifyExpert.asp?tid="+tid);
}
function batchFetchFile(fm) {
	var ids='';
	$(fm).find(':checked[name="sel"]').each(function(index,item){
		ids+=(ids.length?',':'')+item.value;
	});
	tabmgr.goTo('/ThesisReview/admin/batchFetchFile.asp?sel='+ids,'批量下载表格/论文',true);
}
function batchUpdateThesis(fm) {
	$(fm).find(':hidden[name="reviewfilestat"]')
			 .val($('select[name="selreviewfilestat"]').val());
	submitForm(fm,"batchUpdateThesis.asp");
}
function showAllRecords(fm) {
	submitForm(fm,"thesisList.asp?showAll");
}
function rollback(tid,user,opr) {
	if(user!=0&&user!=1&&user!=2&&user!=3) return false;
	var msg=["确实要撤销这个文件的上传操作吗？","确实要撤销这名专家的评阅操作吗？",
					"确实要撤销导师的审核操作吗？","确实要撤销该项操作吗？"];
	if (confirm(msg[user])) {
		submitForm(document.all.fmDetail,"rollback.asp?tid="+tid+"&user="+user+"&opr="+opr);
		return true;
	}
	return false;
}
function modifyReview(tid,rid) {
	submitForm(document.all.fmDetail,"extra/thesisDetail.asp?tid="+tid+"&rev="+rid+"&step=2");
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
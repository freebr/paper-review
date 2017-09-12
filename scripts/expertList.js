function submitForm(fm,action) {
	if(typeof(fm.size)==='function')fm=fm[0];
	if(!!action)fm.action=action;
	fm.submit();
	return false;
}
function showAllRecords(fm) {
	submitForm(fm,"selectExpert.asp?showAll");
}
function showPassword(obj,p) {
	alert('密码为：'+p);
	return;
}
function resetPassword(fm) {
	submitForm(fm,"resetExpertPwd.asp");
}
function batchSendNotice(fm,type) {
	submitForm(fm,"sendmsg.asp?type="+type);
}
function exportInfo(fm) {
	submitForm(fm,"exportExpertInfo.asp");
}
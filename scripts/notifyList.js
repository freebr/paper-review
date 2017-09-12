function submitForm(fm,action) {
	if(typeof(fm.size)==='function')fm=fm[0];
	if(!!action)fm.action=action;
	fm.submit();
	return false;
}
function showNotifyList(fm,usertype) {
	$(fm).find("[name='usertype']").val(usertype);
	submitForm(fm,"notifyList.asp");
}
function notify(fm,utype,uid) {
	submitForm(fm,"notify.asp?sel="+utype+"."+uid);
}
function notifySelection(fm) {
	submitForm(fm,"notify.asp");
}
function showAllRecords(fm) {
	submitForm(fm,"notifyList.asp?showAll");
}
function showTeacherResume(id) {
	window.open("/teacher_resume.asp?id="+id);
	return false;
}
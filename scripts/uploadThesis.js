function checkUploadContent(fm) {
	if(fm.isuploadtable.checked||fm.isuploadthesis.checked)return true;
	alert("请勾选要上传的内容！");
	return false;
}
function submitUploadForm(fm) {
	if(typeof(fm.size)==='function')fm=fm[0];
	fm.btnsubmit.value="正在提交，请稍候……";
	fm.btnsubmit.disabled=true;
	newUploadProgress(fm.uploadid.value,fm.stuid.value);
	return;
}
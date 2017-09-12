function chkForm() {
	$('input.date').each(function() {
		if (!verifyDateTime(this.value)) {
			alert(this.title+"的格式错误！");
			this.focus();
			return false;
		}
	});
	return true;
}
function switchMailContent(n) {
	for (var i=1;i<=11;i++) {
		$('#divmailcontent'+i).css('display',(n==i)?'block':'none');
	}
	return;
}
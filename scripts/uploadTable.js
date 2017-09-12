function verifyDateCh(objVal){
	/* yyyy年mm月dd日 */
	if(objVal.search(/^[0-9]{4}年(0?[1-9]|1[0-2])月(0?[1-9]|1[0-9]|2[0-9]|3[0-1])日$/)==-1)
		return false;
	return true;
}

function submitUploadForm(fm) {
	if(typeof(fm.size)==='function')fm=fm[0];
	var bValid=true;
	$('input.date').each(function() {
		if (!verifyDateCh(this.value)) {
			alert(this.title+"的格式错误！");
			this.focus();
			bValid=false;
		}
	});
	$('input.keyword').each(function() {
		if(!this.dirty) this.value='';
	});
	if(bValid) $(fm.btnsubmit).val("正在提交，请稍候……").attr('disabled',true);
	return bValid;
}

function saveFormDraft() {
	document.cookie=$('form').html();
	return;
}
function readFormDraft() {
	$('form').html(document.cookie);
	return;
}

function initResearchFieldSelectBox(ctl,stu_type) {
	ctl[0].options.length=0;
	$.getJSON('rchfield.asp?type='+stu_type,function(data){
		ctl.data('source',data);
		$.each(data.fields,function(i,elem){
			ctl[0].options.add(new Option(elem.field,i));
		});
		ctl.change();
	});
}

function initSubResearchFieldSelectBox(ctl,ctl_field,field_id) {
	ctl[0].options.length=0;
	$.each(ctl_field.data('source').fields[field_id].sub,function(i,elem){
		ctl[0].options.add(new Option(elem,i));
	});
}

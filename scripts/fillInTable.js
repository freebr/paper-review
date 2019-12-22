function verifyDateCh(objVal){
	/* yyyy年mm月dd日 */
	if(objVal.search(/^[0-9]{4}年(0?[1-9]|1[0-2])月(0?[1-9]|1[0-9]|2[0-9]|3[0-1])日$/)==-1)
		return false;
	return true;
}

function submitUploadForm(fm) {
	var $fm=$(fm);
	if (typeof $fm.form==="function" && !$fm.form('validate')) {
		$.messager.alert("提示", "您填写的信息不完整，请按照提示完善相关内容。", "error");
		return false;
	}
	var bValid=true;
	$('input.date').each(function() {
		if (!verifyDateCh(this.value)) {
			alert(this.title+"的格式错误！");
			this.focus();
			bValid=false;
		}
	});
	if(bValid) $fm.find('#btnsubmit').val("正在提交，请稍候……").attr('disabled',true);
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

function initResearchFieldSelectBox($ctl,stu_type) {
	$ctl[0].options.length=0;
	return $.getJSON('ajax_getResearchField.asp?type='+stu_type,function(data){
		$ctl[0].options.add(new Option('请选择工程领域',''));
		$ctl.data('source',data);
		$.each(data.fields,function(i,elem){
			$ctl[0].options.add(new Option(elem.field,i));
		});
		if(stu_type!=5) $ctl.val('0').parent().hide();
		$ctl.change();
	});
}

function initSubResearchFieldSelectBox($ctl,$ctl_field,field_id,field_text) {
	$ctl[0].options.length=0;
	if(!field_id.length) return;
	field_id=parseInt(field_id);
	$ctl[0].options.add(new Option('请选择研究方向',''));
	var arr=$ctl_field.data('source').fields[field_id].sub;
	$.each(arr,function(i,elem){
		$ctl[0].options.add(new Option(elem,i));
	});
	$ctl[0].options.add(new Option('其他','-1'));
	if (field_text) $ctl.val(arr.indexOf(field_text));
	return;
}
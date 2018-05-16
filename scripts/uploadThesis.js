function checkUploadContent(fm) {
	if(fm.isuploadtable.checked||fm.isuploadthesis.checked) return true;
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

function initAllSubResearchFieldSelectBox($ctl,stu_type,sel_field) {
	$ctl[0].options.length=0;
	$.getJSON('ajax_getResearchField.asp?type='+stu_type,function(data){
		var bFindSelected=false;
		$ctl[0].options.add(new Option('请选择研究方向',''));
		$.each(data.fields,function(i,elem){
			var research_field='【'+elem.field+'领域研究方向】';
			var $option=$(new Option(research_field,'')).addClass('research-field').attr('disabled','disabled');
			$ctl[0].options.add($option[0]);
			$.each(elem.sub,function(j,el) {
				var value=i+'-'+j;
				$ctl[0].options.add(new Option(el,value));
				if(sel_field==el) {
					$ctl.val(value);
					bFindSelected=true;
				}
			});
			$ctl[0].options.add(new Option('其他','-1'));
		});
		!bFindSelected&&$ctl.val('-1');
		$ctl.change();
	});
	return;
}
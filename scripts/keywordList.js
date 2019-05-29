function onKeywordBlur() {
	if(!this.value.length) {
		if(this.name=='keyword_ch')
			this.value='请输入…';
		else
			this.value='Please input...';
		$(this).css({'color':'#999999','font-weight':'bold'});
		this.dirty=false;
	}
}
function onKeywordFocus() {
	if(!this.dirty) {
		this.value='';
		$(this).css({'color':'#000000','font-weight':''});
	}
}
function onKeywordChange() {
	this.dirty=true;
}
function onKeywordRemove() {
	removeKeyword($(this).parents('tr').eq(0));
	return false;
}
function addKeyword() {
	if($('tr.keywordpair').size()>=5) {
		alert('论文关键词不能超过五个！');
		return false;
	}
	var tr=document.createElement('tr');
	var keywordpair=$('tr.keywordpair').eq(-1);
	tr.className='keywordpair';
	tr.innerHTML=keywordpair.html();
	$(tr).find('input.keyword').blur();
	keywordpair.after(tr);
	return;
}
function removeKeyword(keywordpair) {
	if($('tr.keywordpair').size()<=3) {
		alert('论文关键词不能少于三个！');
		return false;
	}
	keywordpair.remove();
	return;
}
function setKeywordCount(new_count) {
	if(new_count<3) {
		alert('论文关键词不能少于三个！');
		return false;
	} else if(new_count>5) {
		alert('论文关键词不能超过五个！');
		return false;
	}
	var count=$('tr.keywordpair').size();
	if(count>new_count) {
		for(var i=0;i<count-new_count;i++) removeKeyword($('tr.keywordpair:last'));
	} else {
		for(var i=0;i<new_count-count;i++) addKeyword();
	}
	return;
}
function setKeywords(keywords_ch,keywords_en) {
	setKeywordCount(keywords_ch.length);
	var i=0;$('input[name="keyword_ch"]').each(function(){onKeywordFocus.call(this);$(this).val(keywords_ch[i++]).change().blur();});
	i=0;$('input[name="keyword_en"]').each(function(){onKeywordFocus.call(this);$(this).val(keywords_en[i++]).change().blur();});
	return;
}
function checkKeywords() {
	var keywords_ch=[], keywords_en=[];
	$.each($.makeArray($('input.keyword')),
		function (index,item){
			if (item.value.indexOf(',')>=0) {
				alert('关键词不能包含半角逗号（,）！');
				item.focus();
				return false;
			}
			if (item.name==='keyword_ch') {
				keywords_ch.push(item.value);
			} else {
				keywords_en.push(item.value);
			}
		}
	);
	$('input[name="keyword_ch"]').val(keywords_ch.join(', '));
	$('input[name="keyword_en"]').val(keywords_en.join(', '));
	return true;
}
$(function() {
	$('input.keyword').on({'blur':onKeywordBlur,'focus':onKeywordFocus,'change':onKeywordChange}).blur();
	$(this).on('click','a.linkAdd',function() {
		addKeyword();
		return false;
	});
	$(this).on('click','a.linkRemove',onKeywordRemove);
	setKeywordCount(3);
});
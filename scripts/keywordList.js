function onKeywordBlur() {
	if(!this.value.length) {
		if(this.name=='keyword_ch')
			this.value='请输入……';
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
	var tr=document.createElement('tr');
	var keywordpair=$('tr.keywordpair').eq(-1);
	tr.className='keywordpair';
	tr.innerHTML=keywordpair.html();
	$(tr).find('a.linkRemove').click(onKeywordRemove);
	$(tr).find('input.keyword').blur(onKeywordBlur).focus(onKeywordFocus).change(onKeywordChange).blur();
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
	var i=0;$('input[name="keyword_ch"]').each(function(){$(this).focus().val(keywords_ch[i++]).change();});
	i=0;$('input[name="keyword_en"]').each(function(){$(this).focus().val(keywords_en[i++]).change();});
	return;
}
function checkKeywords() {
	$('input.keyword').each(function(){
		if(this.value.indexOf(',')>=0) {
			alert('关键词不能包含半角逗号（,）！');
			this.focus();
			return false;
		}
	});
	return true;
}
$().ready(function() {
	$('input.keyword').blur(onKeywordBlur).focus(onKeywordFocus).change(onKeywordChange).blur();
	$('a.linkAdd').click(function() {
		addKeyword();
		return false;
	});
	$('a.linkRemove').click(onKeywordRemove);
	setKeywordCount(3);
});
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
	var i=0;$('input[name="keyword_ch"]').each(function(){$(this).val(keywords_ch[i++]);});
	i=0;$('input[name="keyword_en"]').each(function(){$(this).val(keywords_en[i++]);});
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
			value = item.value.trim();
			if (!value.length) return false;
			if (item.name==='keyword_ch') {
				keywords_ch.push(value);
			} else {
				keywords_en.push(value);
			}
		}
	);
	$('input[name="keyword_ch_all"]').val(keywords_ch.join(', '));
	$('input[name="keyword_en_all"]').val(keywords_en.join(', '));
	return true;
}
$(function() {
	$(this).on('click','a.linkAdd',function() {
		addKeyword();
		return false;
	});
	$(this).on('click','a.linkRemove',onKeywordRemove);
	setKeywordCount(3);
});
$(document).ready(function(){
	$('#btnloaddraft').click(function() {
		var comment_autosaved=getCookie('tutor_comment');
		if(!comment_autosaved) {
			alert('未找到已保存的草稿！');
			return;
		}
		var lastTimeSaved=getCookie('tutor_comment_time');
		if(confirm('将在当前内容结尾处插入上次保存草稿（保存时间为 '+lastTimeSaved+'），是否继续？')) {
			var tb=$('[name="comment"]');
			tb.val(tb.val()+'\r\n\r\n'+comment_autosaved);
		}
	});
});

function verifyDraft(tid) {
	var comment_autosaved=getCookie('tutor_comment');
	var thesis_id_autosaved=getCookie('tutor_thesis_id');
	var lastTimeSaved=getCookie('tutor_comment_time');
	if(comment_autosaved&&thesis_id_autosaved==new String(tid)) {
		if(confirm('系统于 '+lastTimeSaved+' 自动保存了您的草稿，是否读取？')) {
			$('[name="comment"]').val(comment_autosaved).keyup();
		}
	}
}
function saveAsDraft(tid,autosaved) {
	if(!$('[name="comment"]').val().length) return;
	var timeSaved=(new Date()).toLocaleTimeString();
	var tip;
	if(autosaved)
		tip='草稿已自动保存于';
	else
		tip='草稿已保存于';
	setCookie('tutor_thesis_id',tid,24);
	setCookie('tutor_comment',$('[name="comment"]').val(),24);
	setCookie('tutor_comment_time',timeSaved,24);
	$('#tip').html('<font color="blue">'+tip+' '+timeSaved+'</font>');
	return;
}

/*==============================
	Cookie 处理函数
	==============================*/
function setCookie(name,value,expireHours){
	var cookieString=name+"="+escape(value);
	//判断是否设置过期时间
	if(expireHours>0){
		var date=new Date();
		date.setTime(date.getTime+expireHours*3600*1000);
		cookieString=cookieString+"; expire="+date.toGMTString();
	}
	document.cookie=cookieString;
}
function getCookie(name){
	var strCookie=document.cookie;
	var arrCookie=strCookie.split("; ");
	for(var i=0;i<arrCookie.length;i++){
		var arr=arrCookie[i].split("=");
		if(arr[0]==name)return unescape(arr[1]);
	}
	return "";
}
function deleteCookie(name){
	var date=new Date();
	date.setTime(date.getTime()-10000);
	document.cookie=name+"=0; expire="+date.toGMTString();
}
﻿function getFileExt(fn) {
	return fn.substring(fn.lastIndexOf('.')).toLowerCase();
}
function checkIfPdf(ctlupload) {
	if(!ctlupload) return true;
	var fileName=ctlupload.value;
	if (!fileName.length) {
		alert("请为"+ctlupload.title+"选择要上传的 PDF 格式论文文件！");
		return false;
	}
	var fileExt=getFileExt(fileName);
	if (fileExt!=".pdf") {
		alert("所选"+ctlupload.title+"不是 PDF 文件！");
		return false;
	}
	return true;
}
function checkIfWord(ctlupload) {
	if(!ctlupload) return true;
	var fileName=ctlupload.value;
	if (!fileName.length) {
		alert("请为"+ctlupload.title+"选择要上传的 Word 文件！");
		return false;
	}
	var fileExt=getFileExt(fileName);
	if (fileExt!=".doc"&&fileExt!=".docx") {
		alert("所选"+ctlupload.title+"不是 Word 文件！");
		return false;
	}
	return true;
}
function checkIfWordRar(ctlupload) {
	if(!ctlupload) return true;
	var fileName=ctlupload.value;
	if (!fileName.length) {
		alert("请为"+ctlupload.title+"选择要上传的 Word 文件或 RAR 压缩文件！");
		return false;
	}
	var fileExt=getFileExt(fileName);
	if (fileExt!=".doc"&&fileExt!=".docx"&&fileExt!=".rar") {
		alert("所选"+ctlupload.title+"不是 Word 文件或 RAR 压缩文件！");
		return false;
	}
	return true;
}
function createSocket() {
	var sck=new Object();
	try {
		sck.core=new XMLHttpRequest();
	} catch(e) {
		sck.core=new ActiveXObject("MSXML2.XMLHTTP");
	}
	sck.core.onreadystatechange=function(){onData(sck)};
	sck.busy=false;
	return sck;
}
function getProgress(sck,url) {
	if(sck.busy) return;
	sck.busy=true;
	sck.core.open("get","http://"+location.host+'/'+url+"?t="+Number(Math.random(999)*999),true);
	sck.core.send();
}
function newUploadProgress(uploadid,stuid) {
	var div=document.createElement("div");
	div.id="divupload";
	div.className="divupload";
	div.style.width=350;
	div.style.height=80;
	div.style.visibility="visible";
	div.style.left=(parseInt(document.body.offsetWidth)-parseInt(div.style.width))/2+'px';
	div.style.top=(parseInt(document.body.offsetHeight)-parseInt(div.style.height))/2+'px';
	div.innerHTML='<p>正在上传，请稍候……</p>';
	document.body.appendChild(div);
	sckUpload=createSocket();
	sckUpload.onProgress=showUploadProgress;
	setInterval("getProgress(sckUpload,'ThesisReview/student/tmp/"+stuid+uploadid+".json')",1000);
	return;
}
function showUploadProgress() {
	var div=document.getElementById("divupload");
	var perc=uploadProgress.bytesUploaded/uploadProgress.bytesTotal;
	var text="正在上传，请稍候……已完成："+Math.round(perc*10000)/100+'%<br/>'
					+"("+Math.round(uploadProgress.bytesUploaded/102.4)/10+"kB/"+Math.round(uploadProgress.bytesTotal/102.4)/10+"kB)"
	div.innerHTML='<p>'+text+'</p>';
	return;
}
function onData(sck) {
	if(sck.core.readyState==4) {
		if(sck.core.status==200) {
			eval("uploadProgress="+sck.core.responseText);
			sck.onProgress();
			sck.busy=false;
		}
	}
	return;
}
var uploadProgress,sckUpload;
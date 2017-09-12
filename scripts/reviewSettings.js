function getIndexByTeachTypeId(ttid) {
	switch(ttid) {
	case 5:return 0;break;
	case 6:return 1;break;
	case 7:return 2;break;
	case 9:return 3;break;
	}
}
function addReviewTypeItem(rid,typename,teachtypeid,thesisform,reviewfile) {
	var tr,td,tdhtml;
	numItems++;
	maxItems++;
	spanNumItems.innerText=numItems;
	if(!typename) typename='评阅类型'+maxItems+'';
	if(!thesisform) thesisform='';
	tr=document.createElement("tr");
	tr.style.backgroundColor="ghostwhite";
	td=document.createElement("td");
	tdhtml='<span name="identifier">'+numItems+'.</span>&emsp;&emsp;类型名称：<input type="text" name="typename" value="'+typename+'" />'
								 +'&emsp;&emsp;适用专业：<select id="selttid'+numItems+'" name="teachtypeid"><option value="5">ME</option><option value="6">MBA</option><option value="7">EMBA</option><option value="9">MPAcc</option></select>'
								 +'&emsp;&emsp;论文形式名称：<input type="text" name="thesisform" value="'+thesisform+'" /><br/>';
	tdhtml+='&emsp;&emsp;&nbsp;<input type="button" name="removeitem" value="－ 删除" onclick="removeReviewTypeItem(this.parentNode.parentNode)" />';
	if(reviewfile) {
		tdhtml+='&emsp;<a href="upload/review/'+reviewfile+'" target="_blank">查看评阅书</a>';
	}
	tdhtml+='&emsp;上传：<input type="file" name="reviewfile'+numItems+'" accept="application/vnd.ms-word,application/vnd.openxmlformats-officedocument.wordprocessingml.document" /><br/>';
	if(rid) {
		tdhtml+='<input type="hidden" name="rid" value="'+rid+'" />';
	}
	td.height="60px";
	td.innerHTML=tdhtml;
	document.all.tbItems.insertBefore(tr,document.getElementById("trPanel"));
	tr.appendChild(td);
	if(teachtypeid) {
		document.getElementById("selttid"+numItems).selectedIndex=getIndexByTeachTypeId(teachtypeid);
	}
}
function removeReviewTypeItem(tr) {
	if(confirm("是否确定删除本条目？")) {
		document.all.tbItems.removeChild(tr);
		refreshIdentifiers();
	}
}
function refreshIdentifiers() {
	var arr=document.getElementsByName("identifier");
	for(var i=0;i<arr.length;i++)
		arr.item(i).innerText=(i+1)+'.';
	numItems=i;
	spanNumItems.innerText=numItems;
	return;
}
function checkIfPdf(f) {
	var fileName = f.value;
	var fileExt = fileName.substring(fileName.lastIndexOf('.')).toLowerCase();
	if(fileExt.length>0 && fileExt != ".pdf") {
		f.focus();
		alert("所选文件不是 PDF 文件！");
		return false;
	}
	return true;
}
document.all.fmReview.onsubmit=function() {
/*		var inputfile;
	for(var i=0;i<maxItems;i++) {
		inputfile=document.getElementsByName("reviewfile"+(i+1)).item(0);
		if(!inputfile)continue;
		if(!checkIfPdf(inputfile)) {
			return false;
		}
	} */
	this.num_items.value=numItems;
	this.btnadd.disabled=true;
	this.btnsubmit.disabled=true;
}
var numItems=0,maxItems=0;
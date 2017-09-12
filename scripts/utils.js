function vertifyDate(objVal){
	if(objVal.search(/^(\d){4}[-\/](\d){1,2}([-\/](\d){1,2})?$/)==-1)return false;
	arr=objVal.split("-");
	y=parseInt(arr[0]);
	m=parseInt(arr[1].replace(/(^0*)/g, ""));
	if(!!arr[2])
		d=parseInt(arr[2].replace(/(^0*)/g, ""));
	else
		d=1;
	oDate = new Date(y,m-1,d);
	newY=oDate.getYear();
	if(newY<2000)newY+=1900;
	if(y==newY){
    if((m-1)==oDate.getMonth()){
    	if(d==oDate.getDate()){
    		return true;
    	} else {
    		return false;
    	}
    } else {
    	return false;
    }
	} else {
		return false;
	}
}
function vertifyTime(objVal){
	if(objVal.search(/^(\d){1,2}:(\d){1,2}$/)==-1)return false;
	arr=objVal.split(":");
	h=parseInt(arr[0]);
	m=parseInt(arr[1]);
	if(h>23)return false;
	if(m>59)return false;
	return true;
}
function verifyDateTime(objVal){
	/* yyyy-mm-dd || yyyy-mm-dd hh:nn || yyyy-mm-dd hh:nn:ss */
	if(objVal.search(/^[0-9]{4}[-\/](0?[1-9]|1[0-2])[-\/](0?[1-9]|1[0-9]|2[0-9]|3[0-1])( +(0?[0-9]|1[0-9]|2[0-3]):(0?[0-9]|1[0-9]|2[0-9]|3[0-9]|4[0-9]|5[0-9])(:(1[0-9]|2[0-9]|3[0-9]|4[0-9]|5[0-9]|0?[0-9]))?)?$/)==-1)
		return false;
	return true;
}

function compareTime(objVL,objVR){
	arrL=objVL.split(":");
	arrR=objVR.split(":");
	hL=parseInt(arrL[0]);
	mL=parseInt(arrL[1]);
	hR=parseInt(arrR[0]);
	mR=parseInt(arrR[1]);
	result=(hL*60+mL)-(hR*60+mR);
	if(result==0) return 0;//相等
	if(result>0) return 1;//大于
	if(result<0) return 2;//小于
}
function compareDate(objVL,objVR){
    arrL=objVL.split("-");
    arrR=objVR.split("-");
    yL=parseInt(arrL[0]);
    mL=parseInt(arrL[1]);
    dL=parseInt(arrL[2]);
    yR=parseInt(arrR[0]);
    mR=parseInt(arrR[1]);
    dR=parseInt(arrR[2]);
    result=(yL*360+(mL-1)*30+dL)-(yR*360+(mR-1)*30+dR)
    //相等
    if(result==0) return 0;
    //大于
    if(result>0) return 1;
    //小于
    if(result<0) return 2;
}
function isNumber(val){
	if(/[^0123456789]/.test(val)){
	    return false;
	    //找到一个非匹配的模式,就是说不是一个数字
	} else {
	    return true;
	    //找不到一个非匹配的模式,就是说是一个数字
	}
}
function isIDCardNumber(val){
	if(/^(\d{15}|\d{18}|\d{17}x)$/i.test(val))
		return true;
	return false;
}
function isVarchar(val){
	for(i1=0;i1<val.length;i1++)
	{
		if((/[^\w]/.test(val.substring(i1,i1+1))))
		{
		    return false;
		    //找到一个非匹配的模式,就是说不是一个数字+字母
		}
	}
    return true;
    //找不到一个非匹配的模式,就是说是一个数字+字母
}
function isFloat(val){
	//与isFinite 功能相同
	if(/[^0123456789.]/.test(val)){
	    return false;
	    //找到一个非匹配的模式,就是说不是一个数字
	} else {
	    return true;
	    //找不到一个非匹配的模式,就是说是一个数字
	}	    
}

String.prototype.trim = function()    //取空格函数
{
    // 用正则表达式将前后空格
    // 用空字符串替代。
    return this.replace(/(^\s*)|(\s*$)/g, "");
}

String.prototype.trimNoNum = function()    //取非数字函数  	// By Dajing
{
    // 用正则表达式将前后空格
    // 用空字符串替代。
    return this.replace(/[^\d]/g,'');
}

function replNoNum(Obj)	//取非数字函数 与函数trimNoNum共用  	// By Dajing
{
	Obj.value=Obj.value.trimNoNum()
}

function  replClipboardData(Obj)	//在剪切版中 把非数字替换成空 与函数trimNoNum共用	//By Dajing
{
	//clipboardData.setData('text',clipboardData.getData('text').trimNoNum())
	eval("clipboardData.setData('text',clipboardData.getData('text')."+Obj+"())");
}


String.prototype.trimNoFloat = function()    //取非浮点数函数  	// By Dajing
{
    // 用正则表达式将前后空格			
    // 用空字符串替代。
    return this.replace(/[^\d.]/g,'');
}

function replNoFloat(Obj)	//取非浮点数函数 与函数trimNoFloat共用  	// By Dajing
{
	Obj.value=Obj.value.trimNoFloat()
}


function NextFocus(Obj)
{
	//客户跳到下一交点
	if (window.event.keyCode==13)  // 如果是回车
	{
		//Obj 里的值为零时就改变为空
		eval("if (document.all."+Obj+".value=='0') document.all."+Obj+".value='';");
		//设置Obj 为交点
		eval("document.all."+Obj+".focus();");
	}
}


function showCanlendar(objName,col){
	//取日期（完整格式）
 	window.open('../showCanlendar.asp?objName='+objName+'&col='+col,'','width=200,height=240,location=no,scrollbars=no,resizable=no')
}

function showCalendar(objName,sel,fmt){
	//取日期（完整格式）
	url='../../pub/showCalendar.asp?objName='+objName+'&date_sel='+sel;
	if(!!fmt) url+='&fmt='+fmt;
  window.open(url,'','width=200,height=220,location=no,scrollbars=no,resizable=no')
}

function showCleftCanlendar(Cleft,Year,month,col){
		//Cleft 等于 “Cleft” 时取拆分日期 例：2004-11-01 取 year=2004 month=11 (只取年、月)
		// 反之取完整格式 year = date 例：2004-11-01 取 year=2004-11-01		
		if (Cleft=="Cleft")
		{
	        window.open('../showDate.asp?Cleft='+Cleft+'&Year='+Year+'&month='+month+'&col='+col,'','width=200,height=240,location=no,scrollbars=no,resizable=no')
		}
		else
		{
	        window.open('../showDate.asp?Cleft='+Cleft+'&objName='+Year+'&col='+col,'','width=200,height=240,location=no,scrollbars=no,resizable=no')
		}
	
}


function checkEmail( mails )				//判断 EMAIL的格式 是否正确	 与isEmailAddress（）同用			By Dajing
{
	if( mails == "")
		return true;
	mails = mails.replace( /;/g, ",");
	var mailArray = mails.split( ",");
	var unregs = "";
	for( var i=0;i < mailArray.length; i++ )
	{
		var pos = mailArray[i].indexOf( "@");
		if( pos == -1)
		{	unregs = unregs + mailArray[i] + ",\n";
			continue;
		}
		if( mailArray[i].length == pos + 1)
		{	unregs = unregs + mailArray[i] + ",\n";
			continue;
		}
		if( mailArray[i].charAt( mailArray[i].length-1 ) == ">" )
			fix = mailArray[i].substring( pos , mailArray[i].length-1 );
		else
			fix= mailArray[i].substr( pos );
		if( !isEmailAddress(fix)){
			unregs = unregs + mailArray[i] + ",\n";
		}
	}
	if( unregs != "")
	{
		alert( "请正确填写以下Email地址：\n" + unregs.substring( 0, unregs.length-2));
		event.returnValue = false;
		return false;
	}
	return true;
}
function isEmailAddress( email )			//判断 EMAIL的格式 是否正确	 与checkEmail（）同用					By Dajing
{
	var reg=/^@([\w-]+\.)+[\w-]{2,4}$$/;
	return ( email.match( reg ) != null ) ? true:false;
}



function searchObjByTagName(obj, tag)		//从OBJ类中获取TAG 中相应的名字		By Dajing 04.12.20
{
  while(obj!=null && typeof(obj.tagName) != "undefined")
  {
	//获取TAG下所有的OBJECT（类）
    if(obj.tagName == tag.toUpperCase()) return(obj);
    obj = obj.parentElement;			//获取父类 的值集合
  }
  return null;
}

function precision(v, prec)//（四舍五入）取小数点后面几位。  //声哥
{
 var ret=v;
 var ret=v*Math.pow(10,prec);
 ret=Math.round(ret)/Math.pow(10,prec);
 return ret;
}

function isChineseAndNum(s){			//判断输入是否为中文或者数据和字母的函数		//By Dajing
	var ret=true;
	for(var i=0;i<s.length;i++)
	ret=ret && ((s.charCodeAt(i)>=10000)||!(/[^\w]/.test(s.substring(i,i+1))));
	return ret;
}


//设置网页打印的页眉页脚为空 
function PageSetup_Portrait(Obj){ 
	//页面设置 页面方向  
	//	Obj值为Portrait则是纵向
	//	Obj为其它值时则为横向

	var HKEY_Root,HKEY_Path,HKEY_Key,keyWork1,keyWork2; 
	HKEY_Root="HKEY_LOCAL_MACHINE"; 
	HKEY_Path="\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Reliability"; 
	if (Obj='Portrait'){
		 keyWork1="dword:0000260a";
		 keyWork2="hex:d5,07,0c,00,05,00,02,00,03,00,2a,00,0c,00,5a,02,";
	}else{
		 keyWork1="dword:00002592";
		 keyWork2="hex:d5,07,0c,00,05,00,02,00,03,00,28,00,0c,00,58,02,";
	}
	try{ 
		var Wsh=new ActiveXObject("WScript.Shell"); 
		HKEY_Key="LastAliveUptime"; 
		Wsh.RegWrite(HKEY_Root+HKEY_Path+HKEY_Key,keyWork1); 
		HKEY_Key="LastAliveStamp"; 
		Wsh.RegWrite(HKEY_Root+HKEY_Path+HKEY_Key,keyWork2); 
	}catch(e){} 
} 

//开始 对话框函数
function OpenDialog(Action,dHeight,dWidth,dTop,dLeft,reFormName,reValue){	//弹出对话框(打开ASP文件,高 宽 Top与Left的值,返回父窗体FROM名,返回父窗体变量名)	
	var str;
	if (dHeight==""||isNaN(dHeight)){dHeight="450"};
	if (dWidth==""||isNaN(dWidth)){dWidth="400"};
	if (dTop==""||isNaN(dTop)){dTop='100'};
	if (dLeft==""||isNaN(dLeft)){dLeft='300'};

	str="dialogHeight: "+dHeight+"px; dialogWidth: "+dWidth+"px;";
	str=str+" dialogTop: "+dTop+"px; dialogLeft: "+dLeft+"px; edge: Raised;";
	str=str+" center: Yes; help: no; resizable:no; status: no;";
	var ret = window.showModalDialog(Action,'',str);
	if (ret != null&&reValue!="undefined"){
		document.forms[reFormName].all[reValue].value = ret;	
		return true;
	}
	return false;
}
//结束 对话框函数

//开始 层函数
function closeFrame(){
	var oLayerAll=parent.document.getElementsByName("oLayer");
	oLayerAll[oLayerAll.length-1].removeNode(true);
}
function LocateValue(str){
	parent.document.all("aProductName").value = str;
	closeFrame();
}

function getAbsolutePosX(obj) {
	var l = obj.offsetLeft;            
	while(obj = obj.offsetParent)
	l += obj.offsetLeft;
	return l;
}
function getAbsolutePosY(obj) {
	var t = obj.offsetTop;
	while(obj = obj.offsetParent)
	t += obj.offsetTop;
	return t;
}
function ShowDiv(ctrl,action,dWidth,dHeight,dTop,dLeft){		        
	if (dHeight==""||isNaN(dHeight)){dHeight="400"};
	if (dWidth==""||isNaN(dWidth)){dWidth="600"};
	if (dTop==""||isNaN(dTop)){dTop='0'};
	if (dLeft==""||isNaN(dLeft)){dLeft='0'};

	var obj=ctrl;		        
	var leftPos=getAbsolutePosX(obj)+obj.offsetWidth-parseInt(dLeft);
	var topPos=getAbsolutePosY(obj)-parseInt(dTop);

	oDiv=document.createElement("DIV");
	oDiv.id="oLayer";
	oDiv.value="oLayer";
	oDiv.style.cssText="position:absolute;left:"+leftPos+"px;top:"+topPos+"px;width:450px;";
	oDiv.innerHTML="<iframe scrolling='yes' style='width:"+dWidth+"px;height:"+dHeight+"px' frameborder='0' src='"+action+"'></iframe>";
	document.body.appendChild(oDiv);
}
//结束 层函数
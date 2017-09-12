function AddOption(oSelect, Text, Name)
{
	var oOption;
	oOption = document.createElement("OPTION");
	oSelect.options.add(oOption);
	oOption.text = Text;
	oOption.value = Name;
	if (typeof(operatorvalue)=="string"){
		if (operatorvalue==Text)
			oOption.selected=true;
	}
}
function ReloadOperator(){
    var fieldType;
    var fieldName;
    fieldName=document.all.field.options[document.all.field.selectedIndex].value;
    fieldType=fieldName.substring(0,fieldName.indexOf("_"));
    var operator=document.all.operator;
    while (operator.options.length>0)operator.options.remove(0);
    switch(fieldType){
        case 'n':
                AddOption(operator,"大于","大于");
                AddOption(operator,"小于","小于");
                AddOption(operator,"等于","等于");
                AddOption(operator,"不等于","不等于");
                break;
        case 's':case 'ms':
                AddOption(operator,"包含","包含");
                AddOption(operator,"不包含","不包含");
                AddOption(operator,"等于","等于");
                AddOption(operator,"不等于","不等于");
                break;
        case 'd':
                AddOption(operator,"早于","早于");
                AddOption(operator,"迟于","迟于");
                AddOption(operator,"等于","等于");
                AddOption(operator,"不等于","不等于");
                break;
				case 'f':
                AddOption(operator,"等于","等于");
                AddOption(operator,"不等于","不等于");
                break;
        default:
               alert("字段类型不正确!");
        
    }
}
function getFilter(){
    var filter=document.all.filter;
    var filterText="";
    if(filter.value==""){
        filterText="";
        return filterText;
    }
    var fieldType;
    var fieldName,fieldVals;
    var field=document.all.field;
    var filterValue;
    fieldName=field.options[field.selectedIndex].value;
    fieldVal=fieldName.substr(fieldName.indexOf("_")+1);
    fieldType=fieldName.substring(0,fieldName.indexOf("_"));
    var operator=document.all.operator;
    currentOperator=operator.options[operator.selectedIndex].value;
    filterValue=filter.value.replace(/'/g,"''").replace(/\//g,'//').replace(/%/g,'/%').replace(/&/g,'/&');
    switch(fieldType){
        case "n":
                switch(currentOperator){
                    case "大于":filterText=fieldVal+" > "+filterValue;break;
                    case "小于":filterText=fieldVal+" < "+filterValue;break;
                    case "等于":filterText=fieldVal+" = "+filterValue;break;
                    case "不等于":filterText=fieldVal+" <> "+filterValue;break;
                    default:filterText="";
                }
                break;
        case "s":
                switch(currentOperator){
                    case "包含":filterText=fieldVal+" LIKE '%"+filterValue+"%'";break;
                    case "不包含":filterText=fieldVal+" NOT LIKE '%"+filterValue+"%'";break;
                    case "等于":filterText=fieldVal+" = '"+filterValue+"'";break;
                    case "不等于":filterText=fieldVal+" <> '"+filterValue+"'";break;
                    default:filterText="";
                }
                break;
        case "ms":
        				fieldVals=fieldVal.split('|');
        				filterText='';
        				for(var i=0;i<fieldVals.length;i++) {
        					if(!!i) filterText+=' OR ';
	                switch(currentOperator){
	                    case "包含":filterText+=fieldVals[i]+" LIKE '%"+filterValue+"%'";break;
	                    case "不包含":filterText+=fieldVals[i]+" NOT LIKE '%"+filterValue+"%'";break;
	                    case "等于":filterText+=fieldVals[i]+" = '"+filterValue+"'";break;
	                    case "不等于":filterText+=fieldVals[i]+" <> '"+filterValue+"'";break;
	                    default:filterText="";
	                }
              	}
                break;
        case "d":
                switch(currentOperator){
                    case "迟于":filterText=fieldVal+" > '"+filterValue+"'";break;
                    case "早于":filterText=fieldVal+" < '"+filterValue+"'";break;
                    case "等于":filterText=fieldVal+" = '"+filterValue+"'";break;
                    case "不等于":filterText=fieldVal+" <> '"+filterValue+"'";break;
                    default:filterText="";
                }
                break;
				case "f":
                switch(currentOperator){
                    case "等于":filterText=fieldVal+" = '"+filterValue+"'";break;
                    case "不等于":filterText=fieldVal+" <> '"+filterValue+"'";break;
                    default:filterText="";
                }
                break;
        default:filterText="";
        
    }
    if(currentOperator=="包含"||currentOperator=="不包含") filterText+=" ESCAPE '/'";
    return filterText;
    
}
function genFilter(){
    document.all.finalFilter.value=getFilter();
}
function genFinalFilter(formname){
    if(document.all.finalFilter.value!="")
	{
        if(getFilter()!="")document.all.finalFilter.value+=" AND "+getFilter();
	}
    else
	{
    	if(getFilter()!="")document.all.finalFilter.value+=getFilter();
	}
}
function checkAll(){
    idAll=document.getElementsByName("sel");
    for(i=0;i<idAll.length;i++){
        idAll.item(i).checked=document.all.chk.checked;
    }
}
function countClk(){
    var totalCount=0;
    idAll=document.getElementsByName("sel");
    for(i=0;i<idAll.length;i++){
        if(idAll.item(i).checked) totalCount++;
    }
    return totalCount;
}
function chkField(){
	fieldName=document.all.field.options[document.all.field.selectedIndex].value;
	fieldType=fieldName.substring(0,fieldName.indexOf("_"));
	switch(fieldType){
        case "n":
             if(document.all.filter.value!=""){
             	if(!isNumber(document.all.filter.value)){
             		alert("请填写数字!");
             		return false;
             	}
             }
             break;
        //case "s":break;
        case "d":
             if(document.all.filter.value!=""){
             	if(!vertifyDate(document.all.filter.value)){
             		alert("请填写正确的日期格式!\n正确的格式如:2000-1-1");
             		return false;
             	}
             }
             break;
        }
        return true;
}

function checkKey()
{
	if (window.event.keyCode==13) { // 如果是回车
	   genFilter();  //调用东东
	   return true;
	 }
	 return false;
}

function InString(str){
   var str1 = "'"
   var s = str.indexOf(str1);
   return(s);
}

function chkField(){
	fieldName=document.all.field.options[document.all.field.selectedIndex].value;
	fieldType=fieldName.substring(0,fieldName.indexOf("_"));
	switch(fieldType){
        case "s":case "ms":
             if(document.all.filter.value.indexOf("'")!=-1){
                        alert("不能包含单引号(')");
                        return false;
             }
             break;
        case "n":
             if(document.all.filter.value!=""){
             	if(!isNumber(document.all.filter.value)){
             		alert("请填写数字!");
             		return false;
             	}
             }
             break;
        //case "s":break;
        case "d":
             if(document.all.filter.value!=""){
             	if(!vertifyDate(document.all.filter.value)){
             		alert("请填写正确的日期格式!\n正确的格式如:2000-1-1");
             		return false;
             	}
             }
             break;
        }
        return true;
}

function orderFSubmit(Obj,FAction,orderName){
	Obj=document.forms[Obj];
	if (Obj.orderName.value.toUpperCase()==orderName.toUpperCase()){
		Obj.orderName.value=orderName+" desc";
	}else{
		Obj.orderName.value=orderName;
	}
	Obj.action=FAction;
	Obj.submit();
}

function submitForm(form,action) {
	form.action=action;
	form.submit();
	return;
}
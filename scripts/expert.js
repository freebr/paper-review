function submitForm(fm) {
	if(typeof(fm.size)==='function')fm=fm[0];
	with (fm)
	{
		if(teachername.value=="")
		{
			alert("教师姓名不能为空!");
			teachername.focus();
			return false;
		}
		telephone.value=telephone.value.trim();
		if(telephone.value!="")
		{
			if(!isNumber(telephone.value))
			{
				alert("联系电话（办公室）格式有误！");
				telephone.focus();
				return false;
			}
		}
		mobile.value=mobile.value.trim();
		if(mobile.value!="")
		{
			if(!isNumber(mobile.value))
			{
				alert("联系电话（移动）格式有误！");
				mobile.focus();
				return false;
			}
		}
		//判断是否有@字符
		email.value=email.value.trim();
		if(email.value!="")
		{
			if(!checkEmail(email.value))
			{
				email.focus();
				return false;
			}
		}
		//判断身份证号码是否合法
		idcard_no.value=idcard_no.value.trim();
		if(idcard_no.value!="")
		{
			if(!/^\d{17}(\d|X)$/gi.test(idcard_no.value))
			{
				alert("身份证号码格式有误！");
				idcard_no.focus();
				return false;
			}
		}
	  if($(fm).find('#newpwd').size()&&newpwd.value!=repeatpwd.value)
	  {
      alert("新密码和确认新密码不一致!");
      repeatpwd.focus();
      return false;
    }
	}
  fm.submit();
	return true;
}
$(document).ready(function(){
	$('#btnsubmit').click(function(){
		if(!submitForm($('form'))) return;
		$(this).val("正在提交，请稍候...")
					 .attr('disabled',true);
	}).attr('disabled',false);
	$('#btnreturn').click(function(){
		history.go(-1);
	});
});
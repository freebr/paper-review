function showTotalScore() {
	if(!document.powers) return;
	var sum=0;
	var $scores=$(':text[name="scores"]');
	var $scorep=$(':text[name="scorep"]');
	var level,totalValid=true;
	var power1=document.powers.power1;
	var power2=document.powers.power2;
	var i,j,k=0;
	for(var i=0;i<power1.length;i++) {
		var sumPart=0,partValid=true;
		for(var j=0;j<power2[i].length;j++) {
			var s=$scores[k];
			if(!s.value.trim()) {
				totalValid=partValid=false;
				break;
			} else if(isNaN(s.value)) {
				sum="分值无效";
				totalValid=partValid=false;
				break;
			} else if(s.value.indexOf('.')>=0) {
				sum="第"+(i+1)+"项请输入整数";
				totalValid=partValid=false;
				break;
			}
			sumPart+=parseInt(s.value)*power2[i][j];
			k++;
		}
		sumPart*=power1[i];
		if($scorep.size()&&partValid) {
			$scorep[i].value=Math.round(sumPart*100)/100;
		}
		if(totalValid) {
			sum+=sumPart;
		}
	}
	var $total_score=$('span#total_score');
	var $review_level_text=$('span#review_level_text');
	var $review_level=$(':hidden[name="review_level"]');
	var $review_result=$('label[for^="review_result"]');
	if(!totalValid) {
		$review_level_text.html('&nbsp;');
		return;
	}
	sum=Math.round(sum);
	$total_score.text(sum);
	$review_result.hide();
	if(sum<document.remarkStd[2].min) {	// 不合格
		$total_score.css('color','#ff0000');
		$review_level_text.css('color','#ff0000');
		level=document.remarkStd[3].name;i=4;
		$review_result.eq(3).show().find(':radio').attr('checked',true);
	} else {
		$total_score.css('color',"#000000");
		$review_level_text.css('color',"#000000");
		if(sum>=document.remarkStd[0].min) {	// 优秀
			level=document.remarkStd[0].name;i=1;
			$review_result.eq(0).show().find(':radio').attr('checked',true);
		} else if(sum>=document.remarkStd[1].min) {	// 良好
			level=document.remarkStd[1].name;i=2;
			$review_result.eq(0).show().find(':radio').attr('checked',true);
			$review_result.eq(1).show();
		} else {	// 合格
			level=document.remarkStd[2].name;i=3;
			$review_result.eq(1).show().find(':radio').attr('checked',true);
			$review_result.eq(2).show();
		}
	}
	$review_level_text.text(level);
	$review_level.val(i);
	return;
}
function addScoreEventListener() {
	var elems=document.getElementsByName("scores");
	for(var i=0;i<elems.length;i++) {
		elems.item(i).oninput=showTotalScore;
		elems.item(i).onpropertychange=showTotalScore;
	}
	return;
}
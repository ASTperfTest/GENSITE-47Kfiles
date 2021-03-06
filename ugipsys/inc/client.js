//關閉
function btnClose_onclick() {
	window.close();
}

//跳到下一個欄位
function GetFocus(FormId,to) {
  if (to !==null && typeof(to) !=="undefined") {
    var obj=eval(FormId + "." + to);
	if (window.event.keyCode==13)
	{
		//type的屬性是button 就用click的method
		if (obj.type=="button") {
			obj.click();
			//obj.focus();
		} else {
			obj.focus();
		}
	}
  }
}

//刪除字串前後空白 ---
function trim(str) {
	// ^  自字串開頭找起 
	// $  比對字串的結尾
	// \s 符合任何空格之虛格
	// +  比對1 或多次
	str=str.replace(/(^\s+)|(\s+$)|(')/g,"");
	str = str.toUpperCase()
	//將單引號 取代成 兩個單引號
	//str=str.replace(/'/g,"''");
	return str;
}

//取得中英文實際長度
function getLength(str)
{
	var i,cnt=0;
	str = trim(str)
	for(i=0; i<str.length;i++){
		if (escape(str.charAt(i)).length >=4) cnt+=2;
		else cnt++;
	}
	return cnt;
}

//刪除字串間空白, 且轉成大寫
function KillSpace(str)
{
	str = trim(str);
	//去除中間之空格
	str = str.replace(/\s/g,"");
    return str;
}

//檢查身份證字號
function ChkSid(str) {
 /*
 國民身份證統一編號計算公式 :
 一．先將英文字母代號換為數字。A=10
 二．由左至右，第一位乘一，第二位乘九， 第三位乘八， 第四位乘七．．．．．．．
     最後一位乘一。
 三．將各位相對數字所乘之積相加。
 四．將上式(三)所得之和除以十求得餘數。
 五．以十減去上式(四)所得餘數即為檢查號碼。
 */
	var EngStrArr= "ABCDEFGHJKLMNPQRSTUVXYWZIO";
	var total = 0;
	var EngStr,pos;
	EngStr= str.substring(0,1);
	pos= EngStrArr.indexOf(EngStr.toUpperCase());
	//小於10碼
	if((str.length!= 10) || (pos<0)) return false;
	total=parseInt((pos+10)/10,10)*1 + parseInt((pos+10)%10,10)*9
	for(i=1; i<9; i++) {
		total= total + parseInt(str.substring(i,i+1),10)*(9-i);
	}
	total=(10-(total%10))%10;
	if (total==parseInt(str.substring(9,10),10)) {	
		return true;
	} else {
		return false;
	}
}

//聯絡電話檢查 (0-9,-) ---(04-12345678)
function IsPhone(str) {
 str = trim(str);
 //var myReg = /\d{2,3}-\d{7,8}/
 var myReg = new RegExp(/\d{2,3}-\d{7,8}/);
 return (myReg.test(str)) ? true:false;
}

//行動電話檢查 (0-9,-) ---(09xx-123456)
function IsMobilePhone(str) {
 str = trim(str);
 var myReg = new RegExp(/\d{4}-\d{6}/);
 return (myReg.test(str)) ? true:false;
}


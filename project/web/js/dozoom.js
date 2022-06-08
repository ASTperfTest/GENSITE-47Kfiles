//變換字級大小
function changeFontSize(articleObj, act, Status) {

	if (Status=="memfont")
	{
		if (act=="12"){
			act="idx1";
		}
		if (act=="14"){
			act="idx2";
		}	
		if (act=="16"){
			act="idx3";
		}	
		if (act=="18"){
			act="idx4";
	}
	cookieStr=unescape(document.cookie);
    beginStr=cookieStr.indexOf("articleFontSize");
    var articleFontSize
    endStr=cookieStr.indexOf("#N");
	//假如無cookie字型時
    if(beginStr == -1)
    {
	//執行會員字型模式
     articleObj = document.getElementById(articleObj);
    switch (act) {
        case "idx1":
        case "idx2":
        case "idx3":
        case "idx4":
            articleObj.className = act;
            break;
        default:
            switch (articleObj.className) {
                case "idx1":
                    if (act == "max")
                        articleObj.className = "idx2";
                    else
                        articleObj.className = "idx1";
                    break;
                case "idx2":
                    if (act == "max")
                        articleObj.className = "idx3";
                    else
                        articleObj.className = "idx1";
                    break;
                case "idx3":
                    if (act == "max")
                        articleObj.className = "idx4";
                    else
                        articleObj.className = "idx2";
                    break;
                case "idx4":
                    if (act == "max")
                        articleObj.className = "idx4";
                    else
                        articleObj.className = "idx3";
                    break;
            }
            break;
    }
    for (i = 1; i < 5; i++) {
        if (act == "idx" + i) {
            document.getElementById("idx" + i).src = "/subject/images/fontsize_" + i + "_on.gif";
        } else {
            document.getElementById("idx" + i).src = "/subject/images/fontsize_" + i + "_off.gif";
        }
    }
	}
	}
	//Cookie字型模式
	else{
    articleObj = document.getElementById(articleObj);

    //alert (articleObj);
    //document.write(dump_props(articleObj.style,"article"));
    var expires = new Date();
    nowYear = expires.getYear();
    expires.setYear(nowYear + 1);
    expires.setMonth(0);
    expires.setDate(0);
    expires.setHours(0);
    expires.setMinutes(0);
    //	changeFontSize('article','max')
    //	changeFontSize('article','idx3')
    switch (act) {
        case "idx1":
        case "idx2":
        case "idx3":
        case "idx4":
            articleObj.className = act;
            break;
        default:
            switch (articleObj.className) {
                case "idx1":
                    if (act == "max")
                        articleObj.className = "idx2";
                    else
                        articleObj.className = "idx1";
                    break;
                case "idx2":
                    if (act == "max")
                        articleObj.className = "idx3";
                    else
                        articleObj.className = "idx1";
                    break;
                case "idx3":
                    if (act == "max")
                        articleObj.className = "idx4";
                    else
                        articleObj.className = "idx2";
                    break;
                case "idx4":
                    if (act == "max")
                        articleObj.className = "idx4";
                    else
                        articleObj.className = "idx3";
                    break;
            }
            break;
    }
    for (i = 1; i < 5; i++) {
        if (act == "idx" + i) {
            document.getElementById("idx" + i).src = "/subject/images/fontsize_" + i + "_on.gif";
        } else {
            document.getElementById("idx" + i).src = "/subject/images/fontsize_" + i + "_off.gif";
        }
    }
    document.cookie = "articleFontSize=" + articleObj.className + "#N;expires=" + expires.toGMTString() + ";path=/; domain=kwpi-coa-kmweb.gss.com.tw";

}
}
//以下為設定 cookie 的字級
window.onload=function(){
    cookieStr=unescape(document.cookie);
    beginStr=cookieStr.indexOf("articleFontSize");
    var articleFontSize
    endStr=cookieStr.indexOf("#N");
    if(beginStr != -1)
    {	
		//有Cookie字型存在
        articleFontSize=cookieStr.substring((beginStr+16),endStr);
        articleObj=document.getElementById("article").className=articleFontSize;
    }else{
		//無cookie字型存在時使用會員設定的字型
		if (document.getElementById("article").className) {
		articleFontSize=document.getElementById("article").className;
		}
		//無cookie字型存在且無會員設定時的預設值
		else{
        articleFontSize="idx2";
		}
    }
    document.getElementById("idx1").src="/subject/images/fontsize_1_off.gif";
    document.getElementById("idx2").src="/subject/images/fontsize_2_off.gif";
    document.getElementById("idx3").src="/subject/images/fontsize_3_off.gif";
    document.getElementById("idx4").src="/subject/images/fontsize_4_off.gif";
        switch (articleFontSize)
        {
	        case "idx1":
		        var fontSrc="/subject/images/fontsize_1_on.gif";
	        break;
	        case "idx2":
		        var fontSrc="/subject/images/fontsize_2_on.gif";
	        break;
	        case "idx3":
		        var fontSrc="/subject/images/fontsize_3_on.gif";
	        break;
	        case "idx4":
		        var fontSrc="/subject/images/fontsize_4_on.gif";
	        break;
	        default:
		        var fontSrc="/subject/images/fontsize_2_on.gif";
        }
		
        document.getElementById(articleFontSize).src=fontSrc;
};
//以上為設定 cookie 的字級
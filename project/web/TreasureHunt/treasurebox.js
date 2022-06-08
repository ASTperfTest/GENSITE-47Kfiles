var headID = document.getElementsByTagName("head")[0];
var newCss = document.createElement('link');
var treasurehunturl = "/TreasureHunt/treasurehuntservices.aspx";
var mwmberCaptchatUrl = "/TreasureHunt/membercaptchat.aspx";
newCss.type = 'text/css';
newCss.rel = "stylesheet";
newCss.href = "/css/jquery.css";
headID.appendChild(newCss);
var newCss1 = document.createElement('link');
newCss1.type = 'text/css';
newCss1.rel = "stylesheet";
newCss1.href = "/js/greybox/gb_styles.css";
headID.appendChild(newCss1);
var newCss2 = document.createElement('link');
newCss2.type = 'text/css';
newCss2.rel = "stylesheet";
newCss2.href = "/TreasureHunt/include/treasureStyle.css";
headID.appendChild(newCss2);
var dirtyMessage = "";
var isDirty = false;

var GB_ROOT_DIR = "/js/greybox/";


var treasureHtml = "<div id=\"cont\"> ";
treasureHtml += "<div style=\"text-align:right; height:32px\" id=\"treasureclosexx\" ><a onmouseover=this.style.cursor='pointer'; onclick=\"HideTreasureHtml();\"><img src=\"/TreasureHunt/image/close.gif\" width=\"32\" height=\"32\" border=\"0\" /></a></div>";
treasureHtml += "<div id=\"up\"> <table width=\"100%\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">";
treasureHtml += "<tr>";
treasureHtml += "<td align=\"center\"><img src=\"/TreasureHunt/image/{0}\" width=\"207\" height=\"86\" border=\"0\" /></td>";
treasureHtml += "</tr></table></div>";
treasureHtml += "<div id=\"bottom\">";
treasureHtml += "<ul><li class=\"content\" id=\"text_time\" style=\"text-align:center\">{1}</li>";
treasureHtml += "</ul>{2}</div></div>";



function addJs(jsSrc) {
    headID = document.getElementsByTagName("head")[0];
    var newJs = document.createElement('script');
    newJs.type = 'text/javascript';
    newJs.src = jsSrc;
    headID.appendChild(newJs);
}

addJs("/js/greybox/greybox.js");
addJs("/js/greybox/AJS.js");
addJs("/js/greybox/AJS_fx.js");
addJs("/js/greybox/gb_scripts.js");

function doView(baseurl1, title) {
    // var baseurl = ""+baseurl1;

    var h = Math.round(screen.availHeight);
    var w = Math.round(screen.availWidth * 0.8);
    GB_showCenter(title, baseurl1, /* optional */340, 590)
}

treasreLogId = 0;

adWidth = 240; //區塊高度  
TreasureNowX = document.body.clientWidth - 215; //目前位置(left)  
TreasureNowY = 50; //目前位置(top)   

String.format = function () {
    if (arguments.length == 0)
        return null;

    var str = arguments[0];
    for (var i = 1; i < arguments.length; i++) {
        var re = new RegExp('\\{' + (i - 1) + '\\}', 'gm');
        str = str.replace(re, arguments[i]);
    }
    return str;
}

function addDiv(id) {
    // Create the div   
    var div = document.createElement("div");
    // Create the textbox      
    div.id = id;
    div.style.left = 1000 + "px";
    div.style.top = 250 + "px";
    div.style.cssText = "position:absolute;top:100;left:1000;display:none;"
    div.innerHTML = treasureHtml;
    //		div.innerHTML = "<table> " +
    //						"<tr> " +
    //						"<td valign=\"top\" style=\"text-align:center;\" ><div id = \"treasureUp\">" +
    //						"</div><div id=\"text_time\"></div>"+
    //						//ss+
    //						"</td>" + 
    //						"</tr><tr><td><div id=\"treasurebotton\" style=\"text-align:center;\" ></div></td></tr>" +
    //						"</table>"; 
    // Add it all to the form  
    document.body.appendChild(div);
    checkReadDocument();
}

function confirmExit() {
    if (isDirty) {
        return dirtyMessage;
    }
}

window.onbeforeunload = confirmExit;

function checkReadDocument() {
    param = "cmd=checkreaddocument&targetpage=" + location.pathname + "&documentinfo=" + encodeURIComponent(location.search) + "&refereurl=" + referer_url + "&sitecode=" + Math.random()*10;
    jQuery.getJSON(treasurehunturl, param, function (data) {
        if (data != null){
            if (data.Success) {
                if (data.Data.RefererUrl != null && data.Data.RefererUrl != "") {
                    referer_url = data.Data.RefererUrl;
                }
                if (data.Data.Read) {
                    jQuery("#activeFloating").html(String.format(treasureHtml, data.Data.Image, " ", " "));
                    jQuery("#activeFloating").fadeIn(2000);
                    Reciprocalime = data.Data.Reciprocalime;
                    jQuery('#text_time').removeClass("content");
                    jQuery('#text_time').addClass("countdown");
                    document.getElementById('text_time').innerHTML = Reciprocalime;
                    dirtyMessage = decodeURIComponent(data.Data.Message);
                    isDirty = true;
                    showTime();
                    jQuery('#treasureclosexx').html(" ");
                } else if (data.Data) {
                    if (data.Data.Html != null && data.Data.Html != "") {
                        ss = String.format(treasureHtml, data.Data.Image, decodeURIComponent(data.Data.Message), decodeURIComponent(data.Data.Html).replace(/\+/g, " "));
                    } else {
                        ss = String.format(treasureHtml, data.Data.Image, decodeURIComponent(data.Data.Message), " ");
                    }
                    jQuery("#activeFloating").html(ss);

                    jQuery("#activeFloating").fadeIn(2000);
                }
            }
            else {
                jQuery.showAlert(data.Message, data.Data);
				jQuery("#activeFloating").html("");
                return;
            }
			}else{
				jQuery("#activeFloating").html("");
			}
    });
    }
    
function fadIni() {
    TreasureInnerWidth = document.body.clientWidth;
    TreasureInnerHeight = document.body.clientHeight;
    edge = (TreasureInnerWidth - 215);
    if (edge > adWidth) {
        TreasurePosX = edge;
    }
    else {
        TreasurePosX = TreasureInnerWidth - adWidth;
    }

    if (TreasureInnerHeight < 500) {
        TreasurePosY = TreasureInnerHeight - fadHeight;
    }
    else {
        TreasurePosY = 250;
    }
}

function fadRefresh() {
    if (document.body.scrollTop) {
        documentBody = document.body;
    } else {
        documentBody = document.documentElement;
    }
    TreasureOffset = TreasurePosX - TreasureNowX;
    offsetYYYYY = TreasurePosY + documentBody.scrollTop - TreasureNowY;
    TreasureNowX += TreasureOffset / 5;
    TreasureNowY += offsetYYYYY / 5;
    fad_style.left = TreasureNowX + "px";
    fad_style.top = TreasureNowY + "px";
    floatID = setTimeout('fadRefresh()', 20); //Refresh時間 
}
function fadStart() {
    fadIni();
    window.onresize = fadIni;
    fadRefresh();
}

var Reciprocalime = 0;

function showTime() {
    Reciprocalime -= 1;
    document.getElementById('text_time').innerHTML = Reciprocalime;

    if (Reciprocalime == 0) {
        isDirty = false;
        jQuery('#text_time').removeClass("countdown");
        jQuery('#text_time').addClass("content");
        document.getElementById('text_time').innerHTML = "loading...";
        checkTreasureHit();
        //jQuery('#dialog').dialog( 'destroy' );
        //jQuery("#dialog").dialog({
        //	bgiframe: true,
        //	modal: true
        //});
        //jQuery('#dialog').dialog('open');
        //location.href='/SubjectList.aspx';
    }

    //每秒執行一次,showTime()
    if (Reciprocalime != 0)
        setTimeout("showTime()", 1000);
}

function checkTreasureHit() {
    alreadyFinish = true;
    param = "cmd=searchtreasurebox&targetpage=" + location.pathname + "&documentinfo=" + encodeURIComponent(location.search) + "&refereurl=" + referer_url;
    jQuery.getJSON(treasurehunturl, param, function (data) {
        if (data != null)
            if (data.Success) {
                if (data.Data.RefererUrl != null && data.Data.RefererUrl != "")
                    referer_url = data.Data.RefererUrl;
                if (data.Data.Read) {
                    param = "cmd=getmembercaptchat&targetpage=" + location.pathname + "&documentinfo=" + encodeURIComponent(location.search) + "&treasrelogid=" + data.Data.TreasureLodId + "&refereurl=" + referer_url;
                    ss = "<a id=\"treasureCaptchaImage\" onclick=\"changeImage(this)\" ><img name=\"srcCaptchaImage\" src=\"/CaptchaImage/JpegImage.aspx\" alt=\"/CaptchaImage/JpegImage.aspx\"/> </a>";
                    ss += "<input type=\"text\" class=\"txt\" id=\"MemberCaptChaTBox\" size=\"10\" name=\"MemberCaptChaTBox\"/> <input type=\"button\" id=\"MemberCaptCheck\" name = \"MemberCaptCheck\" value=\"Send\" oclick=\"SendMemberCaptChaT();\" /> ";
                    jQuery('#activeFloating').hide();
                    treasreLogId = data.Data.TreasureLodId;
                    doView(mwmberCaptchatUrl + "?" + param, "");
                    //GB_show("123", location.pathname +location.search , /* optional */ 800, 600)
                    //jQuery('#dialog').html(ss);
                    //jQuery('#dialog').dialog('open');
                } else {
                    jQuery("#activeFloating").html(String.format(treasureHtml, data.Data.Image, decodeURIComponent(data.Data.Message), " "));
                }
            }
            else {
                jQuery.showAlert(data.Message, data.Data);
                return;
            }
    });
}

body_obj = 100;

//jQuery(document).ready(function(){
//     jQuery("#dialog").dialog({
//      bgiframe: true, autoOpen: false, height: 100, modal: true
//      ,buttons: {'關閉': function() {  
//           SendMemberCaptChaT(); 
//        }  
//    } 
//    });
//
//});

function changeImage(obj) {

    //alert(document.getElementById("srcCaptchaImage").src);
    //jQuery('#srcCaptchaImage').src('/CaptchaImage/JpegImage.aspx');
    //obj.innerHTML= "/CaptchaImage/JpegImage.aspx";
    //obj.innerHTML = "";
    jQuery.getJSON("/CaptchaImage/JpegImage.aspx", function (data) {
        obj.innerHTML = "<img name=\"srcCaptchaImage\" src=\"/CaptchaImage/JpegImage.aspx\" alt=\"/CaptchaImage/JpegImage.aspx\"/>";
    });
    //document.getElementById("srcCaptchaImage").src = '/CaptchaImage/JpegImage.aspx';
}


function SendMemberCaptChaT() {
    GB_show("123", "", /* optional */800, 600)
    //	    param = "cmd=checkmembrtcaptchat&targetpage="+location.pathname + "&documentinfo=" + encodeURIComponent(location.search) +"&code="+jQuery('#MemberCaptChaTBox').val + "&treasurelodid = "+treasreLogId;
    //	    jQuery.getJSON(url,param,function(data){
    //            if (data.Success) {
    //            if(data.Data){
    //                if(data.Data.Read){
    //                        ss = decodeURIComponent(data.Data.TreasureName);
    //                        jQuery('#dialog').html(ss);
    //                        setTimeout("colseDialog()",100000);
    //                        //jQuery('#dialog').dialog('close');
    //                 }
    //               }else{
    //                jQuery("#text_time").html(decodeURIComponent(data.Message));
    //               }
    //            }
    //            else {
    //                jQuery.showAlert(data.Message, data.Data);
    //                return;
    //            }
    //        });

}

function colseDialog() {
    jQuery('#dialog').dialog('close');
}

function SetGuestNegateTreasure() {
    jQuery('#activeFloating').hide();
    param = "cmd=setguestnegatetreasure";
    jQuery.post(url, param, "");
}

function SetLoginUserNegateTreasure() {
    jQuery('#activeFloating').hide();
    param = "cmd=setloginusernegatetreasure";
    jQuery.post(url, param, "");
}
function HideTreasureHtml() {
    jQuery('#activeFloating').hide();
}
var alreadyFinish = false;

jQuery(document).ready(function () {
    addDiv("activeFloating");
    fad_style = document.getElementById('activeFloating').style;
    fadStart();
});
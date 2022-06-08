<%@ Page Language="C#" AutoEventWireup="true" CodeFile="membercaptchat.aspx.cs" Inherits="TreasureHunt_membercaptchat" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
 <link href="include/treasureStyle.css" rel="stylesheet" type="text/css" />
    <title></title>
</head>
<script src="../js/jquery-1.3.2.min.js" type="text/javascript"></script>
 <script type="text/javascript" language="javascript"> 
    var url = "/TreasureHunt/treasurehuntservices.aspx";
    var treasreLogId= "<%=treasreLogId%>"; 
    var targetpage="<%=targetpage%>";
    var documentinfo="<%=pageParam%>";
    var guid = "<%=guid%>";
    var refereurl = "<%=refereurl %>";
    function SendMemberCaptChaT(){
        $("#SendMember").attr("disabled", true);
        param = "cmd=checkmembrtcaptchat&targetpage=" + targetpage + "&documentinfo=" + encodeURIComponent(documentinfo) + "&code=" + $('#MemberCaptChaTBox').val() + "&treasurelodid=" + treasreLogId + "&guid=" + guid + "&refereurl=" + refereurl;
	    $.getJSON(url,param,function(data){
            if (data.Success) {
            if(data.Data){
                if(data.Data.Read){
                        $('#tres').hide();
                        $('#tres2').show();
                        $('#treasureImage').html("<image src=\"/TreasureHunt/image/big_"+data.Data.Image +"\" class=\"imgbox\" /> ");
                        $('#treasureName').html(decodeURIComponent(data.Data.TreasureName));
                        
                       // ss ="<ul><img src=\"/TreasureHunt/image/"+data.Data.Image +"\" class=\"imgbox\" />";
                        //ss +=" <li class=\"content\">豆仔名稱:</li>";
                        //ss += "<li class=\"content\">您已得到...."+decodeURIComponent(data.Data.TreasureName)+"</li>";
                         //ss += "<li class=\"content\"style=\"text-align:right;\"><input type=\"button\" value=\"關閉視窗\" onclick=\"parent.parent.GB_hide();\" /></li>";
                        //ss += "</ul>";
                        //$('#content').html(ss);
                        //setTimeout("parent.parent.GB_hide();",10000);
                        
                        //$('#dialog').dialog('close');
               }else if(data.Data.Timeout){
                alert(decodeURIComponent(data.Data.Message));
				$('#tres').html(" ");
                parent.parent.GB_hide();
			}
               }else{
                $("#errormessage").html(decodeURIComponent(data.Message));
                $("#SendMember").attr("disabled", false);
               }
            }
            else {
                $.showAlert(data.Message, data.Data);
                return;
            }
        });
	    
	}
	
	function newCode(){
	    var dt = new Date();
		document.getElementById('treasureimage').src="/CaptchaImage/JpegImage.aspx?guid="+guid +"&t=" + dt;	
    }
</script>
<body>
    <form id="form1" runat="server">
        <div id="tres">
        <div id="top"></div>
        <div id="content">
        <ul>
          <li class="content">請輸入驗證碼確認取寶:<a id="A1">
            <img id='treasureimage'  name="srcCaptchaImage" src="/CaptchaImage/JpegImage.aspx?guid=<%=guid%>" alt="/CaptchaImage/JpegImage.aspx" width="131" height="32" /> 
            <input type="button" value="更新圖片" onclick="javascript:newCode(this);" />
        </a> </li>
          
          <li class="content">
            <label>
                <input type="text" class="txt" id="MemberCaptChaTBox" size="10" name="MemberCaptChaTBox"/> 
                </label>
              <label>
                <input type="button" id="SendMember" name = "MemberCaptCheck" value="確認" onclick="SendMemberCaptChaT();" /> 
                </label><label id ="errormessage"></label>
          </li>
          <li class="content2">＊輸入正確的驗證碼確認取寶後，才能於您的百寶箱中
        看到此寶物喔。</li>
            </ul>
        </div>
        </div>
        <div id="tres2" style="display:none;">
            <table width="420"  border="0" align="center" cellpadding="0" cellspacing="0" style="padding: 12px 0 0 0 ">
              <tr>
                <td width="50%" height="240" align="center"><br />
                  <br />      
                  <span id ="treasureImage"></span><br /></td>
                <td width="50%"><span style="font-size:12pt">您獲得的是：</span>
                  <span class="countdown"><br /></span>
                  <span class="treasure" id="treasureName"></span></td>
                </tr>
              <tr>
                <td height="41" colspan="2" align="center"><input type="submit" name="button2" id="button2" value="關閉" onclick="parent.parent.GB_hide();"/></td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>

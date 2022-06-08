<%@ Page Language="C#" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true" CodeFile="index.aspx.cs" Inherits="kmactivity_kmwebpuzzle_index" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>農業知識家推廣活動</title>
  <script type="text/javascript" src="/js/jquery.js"></script>
 <script type="text/javascript" src="/js/swfobject.js"></script>
<link href="include/style.css" rel="stylesheet" type="text/css" />
</head>

<body>
<a title="網頁內容資料" class="Accesskey" accesskey="C" href="#" xmlns:hyweb="urn:gip-hyweb-com">
        :::</a>
        <div class="path" xmlns:hyweb="urn:gip-hyweb-com">
        目前位置：<a title="首頁" href="/mp.asp?mp=1">首頁</a>&gt;<a title="愛拚才會贏" href="#">愛拚才會贏</a></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="puzzle_menu"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="74" align="left" valign="bottom"><a href="01.aspx?a=puzzle"><img src="image/mn_1.gif" width="74" height="28" /></a></td>
        <td width="74" align="left" valign="bottom"><a href="02.aspx?a=puzzle"><img src="image/mn_2.gif" width="74" height="28" /></a></td>
        <td width="74" align="left" valign="bottom"><a href="03.aspx?a=puzzle"><img src="image/mn_3.gif" width="74" height="28" /></a></td>
        <td width="74" align="left" valign="bottom"><a href="07_1.aspx?a=puzzle"><img src="image/mn_7.gif" width="74" height="28" /></a></td>
        <td width="74" align="left" valign="bottom"><a href="gamerank.aspx?a=puzzle"><img src="image/mn_4.gif" width="74" height="28" /></a></td>
        <td width="74" align="left" valign="bottom"><a href="06.aspx?a=puzzle"><img src="image/mn_6.gif" width="74" height="28" /></a></td>
        <td valign="bottom"><a href="#"><img src="image/mn_5h.gif" width="74" height="28" /></a></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td class="content" id="flash_bg">
    <div style=" text-align:center;">
    <script type="text/javascript">
        var flashvars = {};
        var params = {};
        var attributes = {};

        params = {
            autoPlay: "yes"
        };


        if (screen.width > 1024) {
            swfobject.embedSWF("/kmactivity/puzzle/game.swf?token=<%=token %>&url=<%=hosturl%>&v=<%=rr %>", "flashcontent", "600", "540", "9.0.0", "expressInstall.swf", flashvars, params, attributes);
        }
        else if (screen.width == 1024) {
            swfobject.embedSWF("/kmactivity/puzzle/game.swf?token=<%=token %>&url=<%=hosturl%>&v=<%=rr %>", "flashcontent", "600", "540", "9.0.0", "expressInstall.swf", flashvars, params, attributes);

        }
        else {
            swfobject.embedSWF("/kmactivity/puzzle/game.swf?token=<%=token %>&url=<%=hosturl%>&v=<%=rr %>", "flashcontent", "600", "540", "9.0.0", "expressInstall.swf", flashvars, params, attributes);
        }

        function fadRefresh() {
            $.post("/aspSrc.asp", "a=1", "");
            $.post("/subject/aspSrc.asp", "a=1", "");
            $.post("/Pedia/PediaList.aspx", "a=1", "");
            setTimeout('fadRefresh()', 120000); //Refresh時間 
        }
        setTimeout('fadRefresh()', 120000);
</script>


<div id="flashcontent"></div>
        
      
    </div>
</td>
  </tr>
</table>
</body>
</html>
</asp:Content>

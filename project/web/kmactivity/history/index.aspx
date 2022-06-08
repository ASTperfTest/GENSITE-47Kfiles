<%@ Page Language="C#" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true" CodeFile="index.aspx.cs" Inherits="kmactivity_history_index" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
 <html xmlns="http://www.w3.org/1999/xhtml">
 <head id="Head1" runat="server">
    <script type="text/javascript">
        var kmurl = "<%=kmurl %>";
        function hitHint(id) {
            window.open(kmurl + "/Century/Picture_Detail.aspx?ctNodeId=" + id);
            $("#hitNode").val(id);
            $("#hideAnswer").val("");
            $("#aspnetForm").submit();
        }
        function CheckAnswer() {
            $("#answerButton").attr("disabled", true);
            var node = $("input[name='historyanswer']:checked").val();
            if (node == undefined || node == null) {
                alert("請選擇答案!!");
                $("#answerButton").attr("disabled", false);
                return;
            }
            $("#hitNode").val("");
            $("#hideAnswer").val(node);
            $("#aspnetForm").submit();
        }
        function NextQuestionsubmit() {
            $("#hitNode").val("");
            $("#hideAnswer").val("");
            $("#NextQuestion").val("true");
            $("#aspnetForm").submit();
        }
       
        var openPicture = "<%=openPicture %>";
        var canAnswer = "<%=canAnswer %>";
        $(document).ready(function (e) {
            
            if (openPicture == "true") {

            }
            if (canAnswer == "false") {
                $("#nextButton").show();
                $("#hispicarea").show();
                $("#answerButton").hide();
            }
        })
    </script>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>農情濃情拼圖樂</title>
<link href="include/style.css" rel="stylesheet" type="text/css" />
</head>

<body>
<a title="網頁內容資料" class="Accesskey" accesskey="C" href="#" xmlns:hyweb="urn:gip-hyweb-com">
        :::</a>
    <div class="path" xmlns:hyweb="urn:gip-hyweb-com">
        目前位置：<a title="首頁" href="/mp.asp?mp=1">首頁</a>&gt;<a title="農情濃情拼圖樂" href="#">農情濃情拼圖樂</a></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="history_menu"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="80" align="left" valign="bottom"><a href="01.aspx?a=history"><img src="image/mn_1.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="02.aspx?a=history"><img src="image/mn_2.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="03.aspx?a=history"><img src="image/mn_3.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="04.aspx?a=history"><img src="image/mn_7.gif" width="78" height="27" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="top.aspx?a=history"><img src="image/mn_4.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="myspace.aspx?a=history"><img src="image/mn_6.gif" width="78" height="27" /></a></td>
        <td valign="bottom"><a href="#"><img src="image/mn_5h.gif" width="76" height="28" /></a></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td class="content"><div class="content_mid" style="padding:15px 10px; text-align:center">
        <table  border="0" align="center" cellpadding="0" cellspacing="0" id="none">
          <tr>
            <td width="30" height="30" style="background:url(image/frame_tl.gif) no-repeat bottom">&nbsp;</td>
            <td height="106" align="center" style="background:url(image/frame_t.gif) repeat-x bottom">&nbsp;</td>
            <td width="30" style="background:url(image/frame_tr.gif) no-repeat bottom">&nbsp;</td>
          </tr>
          <tr>
            <td style="background:url(image/frame_l.gif) repeat-y right">&nbsp;</td>
            <td style="background:#fff">
                <div align="center" id="hispicarea">
                <image src="<%=imageUrl%>" align="center" style="min-width:200px; min-height:200px;"/>
                </div>
            </td>
            <td style="background:url(image/frame_r.gif) repeat-y left">&nbsp;</td>
          </tr>
          <tr>
            <td height="30" style="background:url(image/frame_bl.gif) no-repeat top">&nbsp;</td>
            <td style="background:url(image/frame_b.gif) repeat-x top">&nbsp;</td>
            <td style="background:url(image/frame_br.gif) no-repeat top">&nbsp;</td>
          </tr>
        </table>
        <div style=" text-align:left"><a href="/Century/Picture_List.aspx?&mp=1" target="_blank">百年農業發展史</a></div><br/>
        <asp:panel id="answerMessageArea" runat="server" visible="false">
            <asp:image id="yesnoimage" runat="server"></asp:image><br/>
        </asp:panel>
        <asp:panel id="tippps" runat="server">
            <h4>請問上圖位於哪一個照片集內</h4><br/>
        </asp:panel>
    
    <div>
        <asp:label runat="server" id="hintstring" text=""></asp:label>
    </div>
    <input type="button" id="answerButton" onclick="CheckAnswer();" value="答題"/>
    <input type="button" id="nextButton" onclick="NextQuestionsubmit();" value="下一題" style=" display:none"/>
    <input type="hidden" value="" id="hitNode" name="hitNode" />
    <input type="hidden" value="" id="hideAnswer" name="hideAnswer" />
    <input type="hidden" value="" id="NextQuestion" name="NextQuestion" />
    <input type="hidden" value="<%=hideGuid %>" id="hideguid" name="hideguid" />
        </div></td>
  </tr>
</table>
</body>

</html>
</asp:Content>

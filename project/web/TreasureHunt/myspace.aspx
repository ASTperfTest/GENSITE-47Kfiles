<%@ Page Language="C#" MasterPageFile="~/4SideMasterPage.master"  AutoEventWireup="true" CodeFile="myspace.aspx.cs" Inherits="TreasureHunt_myspace" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <link href="include/style.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript">
        var buttonDisable = "<%=buttonDisable %>";
        var voteDisplay = "<%=voteDisplay %>";
        jQuery(document).ready(function(){
	     if(buttonDisable=="false"){
	        jQuery("#submitLotteryA").attr("disabled",false);
	        jQuery("#submitLotteryB").attr("disabled",false);
	        jQuery("#submitLotteryC").attr("disabled",false);
	     } 
	     if(voteDisplay == "true"){
	        jQuery("#present").show();
	     }
        });

        function VoteLottery(objectId,voteValue){
            disableButton=  document.getElementById(objectId);
            if(confirm("兌換過後就不可反悔喔!!"))
            {
                disableButton.disabled = (disableButton.disabled == "" ? "disabled" : "");
                document.getElementById("hideVote").value=voteValue;
                document.forms[0].submit();
            }
        }
    </script>
    <a title="網頁內容資料" class="Accesskey" accesskey="C" href="#" xmlns:hyweb="urn:gip-hyweb-com">
        :::</a>
    <div class="path" xmlns:hyweb="urn:gip-hyweb-com">
        目前位置：<a title="首頁" href="/mp.asp?mp=1">首頁</a>&gt;<a title="主題館" href="#">知識尋寶總動員</a></div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="treasuremenu"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="80" align="left" valign="bottom"><a href="01.aspx?avtivityid=2"><img src="image/mn_1.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="02.aspx?avtivityid=2"><img src="image/mn_2.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="03.aspx?avtivityid=2"><img src="image/mn_3.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="04.aspx?avtivityid=2"><img src="image/mn_4.gif" width="78" height="28" /></a></td>
    <td width="68" valign="bottom"><a href="Top.aspx?avtivityid=2"><img src="image/mn_5.gif" width="66" height="28" /></a></td>
    <td valign="bottom"><a href="#"><img src="image/mn_6h.gif" width="107" height="28" /></a></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td class="rank"><div class="content_mid" style="padding:15px 10px; text-align:center">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" id="none">
          <tr>
            <td width="30" height="30" style="background:url(image/qa_br_tl.gif) no-repeat bottom">&nbsp;</td>
            <td align="center" style="background:url(image/qa_br_t.gif) repeat-x bottom"><img src="image/private_h1.gif" width="441" height="180"/></td>
            <td width="30" style="background:url(image/qa_br_tr.gif) no-repeat bottom">&nbsp;</td>
          </tr>
          <tr>
            <td style="background:url(image/qa_br_l.gif) repeat-y right">&nbsp;</td>
            <td style="background:#fff; padding:5px"><table width="100%" border="0" cellpadding="0" cellspacing="0" id="qa">
              <tr>
                <th width="20%" align="center">帳號</th>
                <td width="80%" rowspan="10">
                <asp:label ID="treasureTable" runat="server" text=""></asp:label>
                <asp:panel runat="server">
                  使用兩個寶物隨機換取一個寶物  <asp:DropDownList ID="TreasureDropDownList" AutoPostBack="false" runat="server" /> 
                  <asp:button runat="server" text="交換" id="randomChangeTreasure"  UseSubmitBehavior="false"/>
                </asp:panel>
                 <table width="100%" border="0" cellpadding="0" cellspacing="0" id="present" style="display:none;"> 
                    <tr> 
                      <td colspan="3" align="left"><h2>兌換專區</h2></td> 
                      </tr> 
                    <tr> 
                      <td  align="center"><img src="image/present_a2.jpg" width="80" height="80" /><p>幸運好禮一<br /> 
                          <span class="present_title">25L綠能小冰箱</span></p> 
                          <input type="button" name="submitLotteryA" id="submitLotteryA" onclick="VoteLottery('submitLotteryA','D');"  value="兌換一次抽獎機會" disabled="disabled" class="lotteryButton" > 
                        <p>共<asp:label ID="labelLotteryA" runat="server" text="0" CssClass="chance"></asp:label>次 </p></td> 
                      <td align="center"><img src="image/present_b2.jpg" width="80" height="80" /><p>幸運好禮二<br /> 
                          <span class="present_title">FUJIFILM 拍立得相機</span></p> 
                          <input type="button" name="submitLotteryB" id="submitLotteryB" onclick="VoteLottery('submitLotteryB','E');"  value="兌換一次抽獎機會" disabled="disabled" class="lotteryButton" > 
                        <p>共<asp:label ID="labelLotteryB" runat="server" text="0" CssClass="chance"></asp:label>次 </p></td> 
                      <td  align="center"><img src="image/present_c2.jpg" width="80" height="80" /><p>幸運好禮三<br /> 
                          <span class="present_title">柚花淨白系列</span></p> 
                          <input type="button" name="submitLotteryC" id="submitLotteryC" onclick="VoteLottery('submitLotteryC','F');" value="兌換一次抽獎機會" disabled="disabled" class="lotteryButton" > 
                        <p>共<asp:label ID="labelLotteryC" runat="server" text="0" CssClass="chance"></asp:label>次 </p></td> 
                    </tr> 
                  </table>
                </td>
                </tr>
              <tr align="center">
                <td class="private"><asp:label ID="loginid" runat="server" text=""></asp:label><br /></td>
              </tr>
              <tr align="center">
                <th>暱稱</th>
              </tr>
              <tr align="center">
                <td class="private"><asp:label ID="userName" runat="server" text=""></asp:label></td>
              </tr>
              <tr align="center">
                <th>可抽獎次數</th>
              </tr>
              <tr align="center">
                <td class="private"><asp:label ID="UsersTotalSuit" runat="server" text="0"></asp:label></td>
              </tr>
              <tr align="center">
                <th>已兌換套數</th>
              </tr>
              <tr align="center">
                <td class="private"><asp:label ID="Lottery" runat="server" text="0"></asp:label></td>
              </tr>
              <tr align="center">
                <th>尚未兌換套數</th>
              </tr>
              <tr align="center">
                <td class="private"><asp:label ID="BeforeLottery" runat="server" text="0"></asp:label></td>
              </tr>
            </table></td>
            <td style="background:url(image/qa_br_r.gif) repeat-y left">&nbsp;</td>
          </tr>
          <tr>
            <td height="30" style="background:url(image/qa_br_bl.gif) no-repeat top">&nbsp;</td>
            <td style="background:url(image/qa_br_b.gif) repeat-x top">&nbsp;</td>
            <td style="background:url(image/qa_br_br.gif) no-repeat top">&nbsp;</td>
          </tr>
        </table>
        <input type="hidden" value="" name="hideVote" id="hideVote" />
    </div></td>
  </tr>
</table>
</asp:Content>

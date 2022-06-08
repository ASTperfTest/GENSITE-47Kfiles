<%@ Page Language="C#" MasterPageFile="~/4SideMasterPage.master"  AutoEventWireup="true" CodeFile="06.aspx.cs" Inherits="kmactivity_kmwebpuzzle_05" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>農業知識家推廣活動</title>
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
        <td align="left" valign="bottom"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="74" align="left" valign="bottom"><a href="01.aspx?a=puzzle"><img src="image/mn_1.gif" width="74" height="28" /></a></td>
            <td width="74" align="left" valign="bottom"><a href="02.aspx?a=puzzle"><img src="image/mn_2.gif" width="74" height="28" /></a></td>
            <td width="74" align="left" valign="bottom"><a href="03.aspx?a=puzzle"><img src="image/mn_3.gif" width="74" height="28" /></a></td>
            <td width="74" align="left" valign="bottom"><a href="07_1.aspx?a=puzzle"><img src="image/mn_7.gif" width="74" height="28" /></a></td>
            <td width="74" align="left" valign="bottom"><a href="GameRank.aspx?a=puzzle"><img src="image/mn_4.gif" width="74" height="28" /></a></td>
            <td width="74" align="left" valign="bottom"><a href="06.aspx?a=puzzle"><img src="image/mn_6h.gif" width="74" height="28" /></a></td>
            <td valign="bottom"><a href="index.aspx?a=puzzle"><img src="image/mn_5.gif" width="74" height="28" /></a></td>
          </tr>
        </table></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td class="content"><div class="user_header"><span style="background:url(image/rank_br_t.gif) repeat-x bottom"><img src="image/topimg3.jpg" width="525" height="270"/></span></div>
      <div class="content_mid" style="padding:0 10px 15px 10px; text-align:center">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" id="none">
  <tr>
    <td width="30" height="30" style="background:url(image/rank_br_tl.gif) no-repeat bottom">&nbsp;</td>
    <td height="30" align="center" valign="top" style="background:url(image/rank_br_t.gif) repeat-x bottom">&nbsp;</td>
    <td width="30" style="background:url(image/rank_br_tr.gif) no-repeat bottom">&nbsp;</td>
  </tr>
  <tr>
        <td style="background:url(image/rank_br_l.gif) repeat-y right">&nbsp;</td>
        <td style="background:#fff; padding:5px"><table width="100%" border="0" cellpadding="0" cellspacing="0" id="rank">
          <tr>
                <th >帳號</th>  <th>現有點數</th><th>分數</th> <th>已完成數</th> <th>本日完成數</th><th>本日放棄數</th>
          </tr>
          <tr align="center">
                <td class="private"><asp:label ID="loginid" runat="server" text=""></asp:label><br /></td> 
                 <td class="private"><asp:label ID="userEnergy" runat="server" text="0"></asp:label></td>
                <td class="private"><asp:label ID="UsersTotalScore" runat="server" text="0"></asp:label></td>
                 <td class="private"><asp:label ID="AllQuestion" runat="server" text="0"></asp:label></td>
                 <td class="private"><asp:label ID="UserDailyAnswer" runat="server" text="0"></asp:label></td>
                 <td class="private"><asp:label ID="UserDailyFAnswer" runat="server" text="0"></asp:label></td>
          </tr><tr>
                <td width="80%" colspan="7">
                <br />
                第<asp:Label ID="PageNumberText" runat="server" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server"  CssClass="Number"/>頁，
   	    共<asp:Label ID="TotalRecordText" runat="server" CssClass="Number" />筆資料，     
        <asp:LinkButton ID="preLink" runat="server" Visible="false" OnClick="preLinkAct">
            <asp:Label ID="preText" runat="server" Text="上一頁" />&nbsp;
        </asp:LinkButton>
        跳至第<asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" OnSelectedIndexChanged="ChangePageNumber" runat="server" /> 頁 &nbsp;
        <asp:LinkButton ID="nextLink" runat="server" Visible="false" OnClick="nextLinkAct">
            <asp:Label ID="nextText" runat="server" Text="下一頁" />
        </asp:LinkButton>
                 
                 <asp:Repeater ID="rpList" runat="server">
            <HeaderTemplate>
                <table width="100%" border="0" id="Table1" cellspacing="0" cellpadding="0">
                    <tr>
                        <th >&nbsp;
                        </th>
                        <th >
                            日期
                        </th>
                        <th >
                            完成數
                        </th>
                        <th >
                            放棄數
                        </th>
                        <th>
                            得分
                        </th>
                    </tr>
            </HeaderTemplate>
            <ItemTemplate>
                <tr>
                    <td align="center">
                        <%#Eval("Row") %>
                    </td>
                    <td align="center">
                        <%#(Eval("gametime"))%>
                    </td>
                    <td align="center">
                        <%#Eval("Counting")%>
                    </td>
                    <td align="center">
                        <%#Eval("FCounting")%>
                    </td>
                    <td align="center">
                        <%#Eval("TotalGrade")%>
                    </td>
               </tr>
           </ItemTemplate>
           <FooterTemplate>
               </table>
           </FooterTemplate>
        </asp:Repeater>
                 
               
                </td>
                </tr>
              
            </table></td>
            <td style="background:url(image/rank_br_r.gif) repeat-y left">&nbsp;</td>
          </tr>
          <tr>
            <td height="30" style="background:url(image/rank_br_bl.gif) no-repeat top">&nbsp;</td>
            <td style="background:url(image/rank_br_b.gif) repeat-x top">&nbsp;</td>
            <td style="background:url(image/rank_br_br.gif) no-repeat top">&nbsp;</td>
          </tr>
        </table>
        <input type="hidden" value="" name="hideVote" id="hideVote" />
    </div></td>
  </tr>
</table>
</body>
</html>

</asp:Content>
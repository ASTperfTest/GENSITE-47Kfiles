<%@ Page Language="C#" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true" CodeFile="myspace.aspx.cs" Inherits="kmactivity_history_myspace" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
 <html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <script type="text/javascript">
        $(document).ready(function (e) {
            
        })
    </script>
    <link href="include/style.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <a title="網頁內容資料" class="Accesskey" accesskey="C" href="#" xmlns:hyweb="urn:gip-hyweb-com">
        :::</a>
    <div class="path" xmlns:hyweb="urn:gip-hyweb-com">
        目前位置：<a title="首頁" href="/mp.asp?mp=1">首頁</a>&gt;<a title="農情濃情拼圖樂" href="#">農情濃情拼圖樂</a></div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="treasuremenu"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="80" align="left" valign="bottom"><a href="01.aspx?a=history"><img src="image/mn_1.gif" width="78" height="28" /></a></td>
                    <td width="80" align="left" valign="bottom"><a href="02.aspx?a=history"><img src="image/mn_2.gif" width="78" height="28" /></a></td>
                    <td width="80" align="left" valign="bottom"><a href="03.aspx?a=history"><img src="image/mn_3.gif" width="78" height="28" /></a></td>
                    <td width="80" align="left" valign="bottom"><a href="04.aspx?a=history"><img src="image/mn_7.gif" width="78" height="27" /></a></td>
                    <td width="80" align="left" valign="bottom"><a href="top.aspx?a=history"><img src="image/mn_4.gif" width="78" height="28" /></a></td>
                    <td width="80" align="left" valign="bottom"><a href="#"><img src="image/mn_6h.gif" width="78" height="27" /></a></td>
                    <td valign="bottom"><a href="index.aspx?a=history"><img src="image/mn_5.gif" width="77" height="28" /></a></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td class="content"><div class="content_mid" style="padding:0 10px 15px 10px; text-align:center">
            <table width="100%" border="0" cellpadding="0" cellspacing="0" id="none">
      <tr>
        <td width="30" height="30" style="background:url(image/rank_br_tl.gif) no-repeat bottom">&nbsp;</td>
        <td height="191" align="center" valign="top" style="background:url(image/rank_br_t.gif) repeat-x bottom"><img src="image/user_h1.gif" width="460" height="180"/></td>
        <td width="30" style="background:url(image/rank_br_tr.gif) no-repeat bottom">&nbsp;</td>
      </tr>
      <tr>
        <td style="background:url(image/rank_br_l.gif) repeat-y right">&nbsp;</td>
        <td style="background:#fff; padding:5px"><table width="100%" border="0" cellpadding="0" cellspacing="0" id="rank">
          <tr>
                <th >帳號</th> <th>暱稱</th> <th>分數</th> <th>已答題數</th> <th>答對數</th><th>本日答題數</th>
          </tr>
          <tr align="center">
                <td class="private"><asp:label ID="loginid" runat="server" text=""></asp:label><br /></td> 
                <td class="private"><asp:label ID="userName" runat="server" text=""></asp:label></td>
                <td class="private"><asp:label ID="UsersTotalScore" runat="server" text="0"></asp:label></td>
                 <td class="private"><asp:label ID="AllQuestion" runat="server" text="0"></asp:label></td>
                 <td class="private"><asp:label ID="CurrentQuestion" runat="server" text="0"></asp:label></td>
                 <td class="private"><asp:label ID="UserDailyAnswer" runat="server" text="0"></asp:label></td>
          </tr><tr>
                <td width="80%" colspan="6">
                <br />
                 第<asp:Label ID="PageNumberText" runat="server" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server"  CssClass="Number"/>頁，
   	                    共<asp:Label ID="TotalRecordText" runat="server" CssClass="Number" />筆資料，     
                        <asp:HyperLink ID="PreviousLink" runat="server">
                          <!--<asp:image runat="server" ImageUrl="/xslGip/style1/images3/arrow_left.gif" ID="PreviousImg" AlternateText="上一頁" /> -->
                          <asp:Label ID="PreviousText" runat="server" >上一頁 &nbsp;</asp:Label>
                        </asp:HyperLink>
                        跳至第<asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" /> 頁 &nbsp;
                        <asp:HyperLink ID="NextLink" runat="server">
                          <!--<asp:image runat="server" ImageUrl="/xslGip/style1/images3/arrow_right.gif" ID="NextImg" AlternateText="下一頁"/> -->
                          <asp:Label ID="NextText" runat="server">下一頁 &nbsp;</asp:Label>
                        </asp:HyperLink>
                 
                 <asp:label ID="treasureTable" runat="server" text=""></asp:label>
                 
               
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
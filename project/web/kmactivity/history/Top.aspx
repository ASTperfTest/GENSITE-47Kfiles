<%@ Page Language="C#"  MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true" CodeFile="Top.aspx.cs" Inherits="kmactivity_history_Top" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <link href="include/style.css" rel="stylesheet" type="text/css" />
    <a title="網頁內容資料" class="Accesskey" accesskey="C" href="#" xmlns:hyweb="urn:gip-hyweb-com">
        :::</a>
        <div class="path" xmlns:hyweb="urn:gip-hyweb-com">
        目前位置：<a title="首頁" href="/mp.asp?mp=1">首頁</a>&gt;<a title="農情濃情拼圖樂" href="#">農情濃情拼圖樂</a></div>
    <div>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td class="treasuremenu"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="80" align="left" valign="bottom"><a href="01.aspx?a=history"><img src="image/mn_1.gif" width="78" height="28" /></a></td>
                    <td width="80" align="left" valign="bottom"><a href="02.aspx?a=history"><img src="image/mn_2.gif" width="78" height="28" /></a></td>
                    <td width="80" align="left" valign="bottom"><a href="03.aspx?a=history"><img src="image/mn_3.gif" width="78" height="28" /></a></td>
                    <td width="80" align="left" valign="bottom"><a href="04.aspx?a=history"><img src="image/mn_7.gif" width="78" height="27" /></a></td>
                    <td width="80" align="left" valign="bottom"><a href="#"><img src="image/mn_4h.gif" width="78" height="28" /></a></td>
                    <td width="80" align="left" valign="bottom"><a href="myspace.aspx?a=history"><img src="image/mn_6.gif" width="78" height="27" /></a></td>
                    <td valign="bottom"><a href="index.aspx?a=history"><img src="image/mn_5.gif" width="77" height="28" /></a></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td class="content"><div class="content_mid" style="padding:15px 10px; text-align:center">
                    <table width="100%" border="0" cellpadding="0" cellspacing="0" id="none">
                      <tr>
                        <td width="30" height="30" style="background:url(image/rank_br_tl.gif) no-repeat bottom">&nbsp;</td>
                        <td height="191" align="center" valign="top" style="background:url(image/rank_br_t.gif) repeat-x bottom"><img src="image/rank_h1.gif" width="460" height="180"/></td>
                        <td width="30" style="background:url(image/rank_br_tr.gif) no-repeat bottom">&nbsp;</td>
                      </tr>
                      <tr>
                        <td style="background:url(image/rank_br_l.gif) repeat-y right">&nbsp;</td>
                          <td style="background:#fff; padding:5px">
                            <asp:Label ID="TableText" runat="server" Text="" />
                        </td>
                        <td style="background:url(image/rank_br_r.gif) repeat-y left">&nbsp;</td>
                      </tr>
                      <tr>
                         <td height="30" style="background:url(image/rank_br_bl.gif) no-repeat top">&nbsp;</td>
                        <td style="background:url(image/rank_br_b.gif) repeat-x top">&nbsp;</td>
                        <td style="background:url(image/rank_br_br.gif) no-repeat top">&nbsp;</td>
                      </tr>
                    </table>
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
                        </asp:HyperLink>，每頁                      
                        <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
                          <asp:ListItem Selected="True" Value="15">15</asp:ListItem>                    
                          <asp:ListItem Value="30">30</asp:ListItem>
                          <asp:ListItem Value="50">50</asp:ListItem>
                        </asp:DropDownList>筆資料
                     </div></td>
              </tr>
            </table>
      
      <hr />    
      <div class="top">
        <a href="#" title="top">top</a>
      </div>
    </div>
</asp:Content>

<%@ Page Language="C#"  MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true" CodeFile="GameRank.aspx.cs" Inherits="GameRank" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <link href="include/style.css" rel="stylesheet" type="text/css" />
    <a title="�������e���" class="Accesskey" accesskey="C" href="#" xmlns:hyweb="urn:gip-hyweb-com">
        :::</a>
        <div class="path" xmlns:hyweb="urn:gip-hyweb-com">
        �ثe��m�G<a title="����" href="/mp.asp?mp=1">����</a>&gt;<a title="�R��~�|Ĺ" href="#">�R��~�|Ĺ</a></div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="puzzle_puzzle_menu"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td align="left" valign="bottom"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="74" align="left" valign="bottom"><a href="01.aspx?a=puzzle"><img src="image/mn_1.gif" width="74" height="28" /></a></td>
            <td width="74" align="left" valign="bottom"><a href="02.aspx?a=puzzle"><img src="image/mn_2.gif" width="74" height="28" /></a></td>
            <td width="74" align="left" valign="bottom"><a href="03.aspx?a=puzzle"><img src="image/mn_3.gif" width="74" height="28" /></a></td>
            <td width="74" align="left" valign="bottom"><a href="07_1.aspx?a=puzzle"><img src="image/mn_7.gif" width="74" height="28" /></a></td>
            <td width="74" align="left" valign="bottom"><a href="#"><img src="image/mn_4h.gif" width="74" height="28" /></a></td>
            <td width="74" align="left" valign="bottom"><a href="06.aspx?a=puzzle"><img src="image/mn_6.gif" width="74" height="28" /></a></td>
            <td valign="bottom"><a href="index.aspx?a=puzzle"><img src="image/mn_5.gif" width="74" height="28" /></a></td>
          </tr>
        </table></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td class="content"><div class="content_header"><span style="background:url(image/rank_br_t.gif) repeat-x bottom"><img src="image/topimg2.jpg" width="525" height="270" /></span></div>
      <div class="content_mid" style="padding:0 10px 15px 10px; text-align:center">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" id="none">
  <tr>
    <td width="30" height="30" style="background:url(image/rank_br_tl.gif) no-repeat bottom">&nbsp;</td>
    <td height="30" align="center" valign="top" style="background:url(image/rank_br_t.gif) repeat-x bottom">&nbsp;</td>
    <td width="30" style="background:url(image/rank_br_tr.gif) no-repeat bottom">&nbsp;</td>
  </tr>
  <tr>
    <td style="background:url(image/rank_br_l.gif) repeat-y right">&nbsp;</td>
    <td style="background:#fff; padding:5px">
    <!--�����˦����٨S�n��!!!!-->
      <asp:Repeater ID="rpList" runat="server">
            <HeaderTemplate>
                <table width="100%" border="0" id="rank" cellspacing="0" cellpadding="0">
                    <tr>
                        <th >&nbsp;
                        </th>
                        <th >
                            �m�W�x�ʺ�
                        </th>
                        <th >
                            ���ʱo��
                        </th>
                        <th>
                            ������
                        </th>
                        <th>
                            �ϥ��I��
                        </th>
                    </tr>
            </HeaderTemplate>
            <ItemTemplate>
                <tr>
                    <td align="center">
                        <%#Eval("Row") %>
                    </td>
                    <td align="center">
                        <%#DealName(Eval("REALNAME"), Eval("NICKNAME"))%>
                    </td>
                    <td align="center">
                        <%#Eval("TotalGrade")%>
                    </td>
                    <td align="center">
                        <%#Eval("Counting")%>
                    </td>
                    <td align="center">
                        <%#Eval("useenergy")%>
                    </td>
               </tr>
           </ItemTemplate>
           <FooterTemplate>
               </table>
           </FooterTemplate>
        </asp:Repeater>
        </td>
    <td style="background:url(image/rank_br_r.gif) repeat-y left">&nbsp;</td>
  </tr>
  <tr>
    <td height="30" style="background:url(image/rank_br_bl.gif) no-repeat top">&nbsp;</td>
    <td style="background:url(image/rank_br_b.gif) repeat-x top">&nbsp;</td>
    <td style="background:url(image/rank_br_br.gif) no-repeat top">&nbsp;</td>
  </tr>
</table>
        ��<asp:Label ID="PageNumberText" runat="server" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server"  CssClass="Number"/>���A
   	    �@<asp:Label ID="TotalRecordText" runat="server" CssClass="Number" />����ơA     
        <asp:LinkButton ID="preLink" runat="server" Visible="false" OnClick="preLinkAct">
            <asp:Label ID="preText" runat="server" Text="�W�@��" />&nbsp;
        </asp:LinkButton>
        ���ܲ�<asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" OnSelectedIndexChanged="ChangePageNumber" runat="server" /> �� &nbsp;
        <asp:LinkButton ID="nextLink" runat="server" Visible="false" OnClick="nextLinkAct">
            <asp:Label ID="nextText" runat="server" Text="�U�@��" />
        </asp:LinkButton>
        �A�C��                      
        <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" OnSelectedIndexChanged="ChangePageSize" runat="server">
        <asp:ListItem Selected="True" Value="15">15</asp:ListItem>                    
        <asp:ListItem Value="30">30</asp:ListItem>
        <asp:ListItem Value="50">50</asp:ListItem>
        </asp:DropDownList>�����
      <div class="top">
        <a href="#" title="top">top</a>
      </div>
      </td>
  </tr>
</table>
</asp:Content>

<%@ Page Language="C#" MasterPageFile="~/4SideMasterPage.master"  AutoEventWireup="true" CodeFile="06.aspx.cs" Inherits="kmactivity_kmwebpuzzle_05" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>�A�~���Ѯa���s����</title>
<link href="include/style.css" rel="stylesheet" type="text/css" />
</head>

<body>
 <a title="�������e���" class="Accesskey" accesskey="C" href="#" xmlns:hyweb="urn:gip-hyweb-com">
        :::</a>
        <div class="path" xmlns:hyweb="urn:gip-hyweb-com">
        �ثe��m�G<a title="����" href="/mp.asp?mp=1">����</a>&gt;<a title="�R��~�|Ĺ" href="#">�R��~�|Ĺ</a></div>
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
                <th >�b��</th>  <th>�{���I��</th><th>����</th> <th>�w������</th> <th>���駹����</th><th>�������</th>
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
                ��<asp:Label ID="PageNumberText" runat="server" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server"  CssClass="Number"/>���A
   	    �@<asp:Label ID="TotalRecordText" runat="server" CssClass="Number" />����ơA     
        <asp:LinkButton ID="preLink" runat="server" Visible="false" OnClick="preLinkAct">
            <asp:Label ID="preText" runat="server" Text="�W�@��" />&nbsp;
        </asp:LinkButton>
        ���ܲ�<asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" OnSelectedIndexChanged="ChangePageNumber" runat="server" /> �� &nbsp;
        <asp:LinkButton ID="nextLink" runat="server" Visible="false" OnClick="nextLinkAct">
            <asp:Label ID="nextText" runat="server" Text="�U�@��" />
        </asp:LinkButton>
                 
                 <asp:Repeater ID="rpList" runat="server">
            <HeaderTemplate>
                <table width="100%" border="0" id="Table1" cellspacing="0" cellpadding="0">
                    <tr>
                        <th >&nbsp;
                        </th>
                        <th >
                            ���
                        </th>
                        <th >
                            ������
                        </th>
                        <th >
                            ����
                        </th>
                        <th>
                            �o��
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
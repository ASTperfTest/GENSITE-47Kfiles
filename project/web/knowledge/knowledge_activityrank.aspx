<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="knowledge_activityrank.aspx.vb" Inherits="knowledge_activityrank" title="活動排行榜" %>

<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <link href="include/style.css" rel="stylesheet" type="text/css" />
	<uc1:LeftMenu ID="LeftMenu1" runat="server" />
  
	<!--中間內容區--> 
	<td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

		<div class="path">目前位置：<a href="/">首頁</a>&gt;<a href="knowledge_activityrank.aspx">排行榜</a></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td class="menuA"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_introduce.aspx"><img src="image/mn_1.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_rule.aspx"><img src="image/mn_2.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_statement.aspx"><img src="image/mn_3.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/knowledge_activityrank.aspx"><img src="image/mn_4h.gif" width="78" height="28" /></a></td>
        <td valign="bottom"><asp:Label ID="TabText" runat="server" Text=""></asp:Label></td>
      </tr>
    </table></td>
    <td class="menuA">&nbsp;</td>
    <td class="menuA">&nbsp;</td>
    <td class="menuA">&nbsp;</td>
    <td class="menuA">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="5" class="content">
	
      <div class="content_mid" style="padding:15px 10px; text-align:center">
		 <table width="100%" border="0" cellpadding="0" cellspacing="0" id="none">
		  <tr>
			<td width="30" height="30" style="background:url(image/rank_br_tl.gif) no-repeat bottom">&nbsp;</td>
			<td align="left" style="background:url(image/rank_br_t.gif) repeat-x bottom"><img src="image/rank_h1.gif" width="262" height="133"/></td>
			<td width="100" style="background:url(image/rank_br_t.gif) repeat-x bottom"><a href="/knowledge/myknowledge_record.aspx" class="link">查詢我的積分</a></td>
			<td width="30" style="background:url(image/rank_br_tr.gif) no-repeat bottom">&nbsp;</td>
		  </tr>
		  <tr>
			<td style="background:url(image/rank_br_l.gif) repeat-y right">&nbsp;</td>
			<td colspan="2" style="background:#fff; padding:5px">
				<asp:Label ID="TableText" runat="server" Text="" />   
				</td>
    <td style="background:url(image/rank_br_r.gif) repeat-y left">&nbsp;</td>
  </tr>
  <tr>
    <td height="30" style="background:url(image/rank_br_bl.gif) no-repeat top">&nbsp;</td>
    <td colspan="2" style="background:url(image/rank_br_b.gif) repeat-x top">&nbsp;</td>
    <td style="background:url(image/rank_br_br.gif) no-repeat top">&nbsp;</td>
  </tr>
</table>
			  <div class="Page">
				第<asp:Label ID="PageNumberText" runat="server" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server"  CssClass="Number"/>頁，
				共<asp:Label ID="TotalRecordText" runat="server" CssClass="Number" />筆資料，     
				<asp:HyperLink ID="PreviousLink" runat="server">
				  <asp:image runat="server" ImageUrl="/xslGip/style1/images3/arrow_left.gif" ID="PreviousImg" AlternateText="上一頁" />
				  <asp:Label ID="PreviousText" runat="server" >上一頁 &nbsp;</asp:Label>
				</asp:HyperLink>
				跳至第<asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" /> 頁 &nbsp;
				<asp:HyperLink ID="NextLink" runat="server">
				  <asp:image runat="server" ImageUrl="/xslGip/style1/images3/arrow_right.gif" ID="NextImg" AlternateText="下一頁"/>
				  <asp:Label ID="NextText" runat="server">下一頁 &nbsp;</asp:Label>
				</asp:HyperLink>，每頁                      
				<asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
				  <asp:ListItem Selected="True" Value="15">15</asp:ListItem>                    
				  <asp:ListItem Value="30">30</asp:ListItem>
				  <asp:ListItem Value="50">50</asp:ListItem>
				</asp:DropDownList>筆資料
			  </div>

				
		    </div>
			</p>
		</div></td>
  </tr>
</table>
		
		
	  <div class="top">
			<a href="#" title="top">top</a>
	  </div>   
			
  </td> 
</asp:Content>


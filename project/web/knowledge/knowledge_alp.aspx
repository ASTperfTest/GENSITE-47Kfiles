<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="knowledge_alp.aspx.vb" Inherits="knowledge_alp" title="農業知識入口網-農業知識家" ValidateRequest="false" %>

<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <link href="include/style.css" rel="stylesheet" type="text/css" />
  
  <uc1:LeftMenu ID="LeftMenu1" runat="server" />
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

	  <div class="path">目前位置：<a href="/knowledge/knowledge.aspx">首頁</a><asp:Label ID="PathText" runat="server" Text=""></asp:Label></div>
	 
	 <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="menuA"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_introduce.aspx"><img src="image/mn_1.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_rule.aspx"><img src="image/mn_2.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_statement.aspx"><img src="image/mn_3.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/knowledge_activityrank.aspx"><img src="image/mn_4.gif" width="78" height="28" /></a></td>
        <td valign="bottom"><a href="/knowledge/knowledge_alp.aspx"><img src="image/mn_5h.gif" width="107" height="28" /></a></td>
      </tr>
    </table></td>
    <td class="menuA">&nbsp;</td>
    <td class="menuA">&nbsp;</td>
    <td class="menuA">&nbsp;</td>
    <td class="menuA">&nbsp;</td>
  </tr>
  <tr>
    <td class="content">
      <div class="content_mid"  style="padding:15px 10px; text-align:center">
      <table width="100%" border="0" cellpadding="0" cellspacing="0" id="none">
          <tr>
            <td width="30" height="30" style="background:url(image/qa_br_tl.gif) no-repeat bottom">&nbsp;</td>
            <td align="center" style="background:url(image/qa_br_t.gif) repeat-x bottom"><img src="image/qa_h1.gif" width="432" height="104"/></td>
            <td width="30" style="background:url(image/qa_br_tr.gif) no-repeat bottom">&nbsp;</td>
          </tr>
          <tr>
            <td style="background:url(image/qa_br_l.gif) repeat-y right">&nbsp;</td>
            <td style="background:#fff; padding:5px">
			<asp:Label ID="TableText" runat="server" Text="" />  
			</td>
            <td style="background:url(image/qa_br_r.gif) repeat-y left">&nbsp;</td>
          </tr>
          <tr>
            <td height="30" style="background:url(image/qa_br_bl.gif) no-repeat top">&nbsp;</td>
            <td style="background:url(image/qa_br_b.gif) repeat-x top">&nbsp;</td>
            <td style="background:url(image/qa_br_br.gif) no-repeat top">&nbsp;</td>
          </tr>
        </table>
		
		<div class="Event">			    
		    <div class="Page">
        	    第<asp:Label ID="PageNumberText" runat="server" Text="" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server" Text="" CssClass="Number" />頁，
        	    共<asp:Label ID="TotalRecordText" runat="server" Text="" CssClass="Number" />筆資料，     
        	    <asp:HyperLink ID="PreviousLink" runat="server">
        	        <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg" Visible="false" AlternateText="上一頁"></asp:image>        	            
        	        <asp:Label ID="PreviousText" runat="server" Visible="false" Text="Label">上一頁 &nbsp;</asp:Label>            
        	    </asp:HyperLink>
                到第 <asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" /> 頁 &nbsp;
                <asp:HyperLink ID="NextLink" runat="server">
                    <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg" Visible="false" AlternateText="下一頁"></asp:image>                     	
                    <asp:Label ID="NextText" runat="server" Visible="false" Text="Label">下一頁 &nbsp;</asp:Label>
                </asp:HyperLink>，每頁顯示                      
                <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
                    <asp:ListItem Selected="True" Value="10">10</asp:ListItem>
                    <asp:ListItem Value="20">20</asp:ListItem>
                    <asp:ListItem Value="30">30</asp:ListItem>
                    <asp:ListItem Value="50">50</asp:ListItem>
                </asp:DropDownList> 筆
            </div>
              
	    </div>
		</div></td>
  </tr>
</table>
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td> 
</asp:Content>


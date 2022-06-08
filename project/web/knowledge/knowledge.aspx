<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="knowledge.aspx.vb" Inherits="knowledge" title="農業知識入口網-農業知識家" ValidateRequest="false" %>
 
<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <uc1:LeftMenu ID="LeftMenu1" runat="server" />
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />
            
    <!--主題推薦-->
    <asp:Label ID="TopicRecommandText" runat="server" Text=""></asp:Label>
	
    <!--最新發問-->
    <asp:Label ID="LatestQuestionText" runat="server" Text=""></asp:Label>
    
    <!--熱門討論-->
    <asp:Label ID="HotDiscussText" runat="server" Text=""></asp:Label>
    
    <!--專家補充-->
    <asp:Label ID="ProfessionalText" runat="server" Text=""></asp:Label>

		<div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->

  </td>
  
</asp:Content>


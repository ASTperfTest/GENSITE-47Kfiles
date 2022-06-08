<%@ Control Language="VB" AutoEventWireup="false" CodeFile="LeftMenu.ascx.vb" Inherits="knowledge_LeftMenu" %>

  <td class="left">
      
    <!--左方menu 知識統計-->
    <div class="sta"><asp:Label ID="KnowledgeSummaryText" runat="server" Text=""/></div>            

    <!--左方menu 農業知識分類-->
    <ul class="menu2"><asp:Label ID="KnowledgeTypeLink" runat="server" Text=""/></ul>            

    <!--左方menu RSS-->
    <div class="rss"><a href="/lp.asp?ctNode=1574&CtUnit=305&BaseDSD=7&mp=1">RSS訂閱服務</a></div>            

    <!--左方menu 廣告-->
    <div class="ad"><ul><asp:Label ID="AdRotateLink" runat="server" Text=""/></ul></div>
    
  </td>

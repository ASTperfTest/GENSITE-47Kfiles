<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="knowledge_lp.aspx.vb" Inherits="knowledge_lp" title="農業知識入口網-農業知識家" ValidateRequest="false" %>

<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <uc1:LeftMenu ID="LeftMenu1" runat="server" />
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

	  <div class="path">目前位置：<a href="/knowledge/knowledge.aspx">首頁</a><asp:Label ID="PathText" runat="server" Text=""></asp:Label></div>
	  <!--div class="hotQa"><ul class="list"><li>重要提醒!! 目前知識一起答活動正在進行中，想挑戰活動之會員，請 <a href="/knowledge/knowledge_alp.aspx">直接點選這裡瀏覽本活動範圍之問題</a><IMG src="http://kmbeta.coa.gov.tw/xslGip/style1/images2/good.gif"></li></ul></div-->
	  <div id="Magazine">
		<div id="MagTabs"><asp:Label ID="TabText" runat="server" Text=""></asp:Label></div>		
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
            <asp:Label ID="TableText" runat="server" Text="" />    
	    </div>
	  </div>
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td> 
</asp:Content>


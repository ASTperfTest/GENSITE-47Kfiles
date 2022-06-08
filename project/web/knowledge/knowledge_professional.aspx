<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="knowledge_professional.aspx.vb" Inherits="knowledge_professional" title="農業知識入口網-農業知識家" ValidateRequest="false" %>

<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <script src="http://www.google-analytics.com/urchin.js" type="text/javascript"></script>

  <uc1:LeftMenu ID="LeftMenu1" runat="server" />
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />
	
	  <div class="path">目前位置：<a href="/knowledge/knowledge.aspx">首頁</a><asp:Label ID="PathText" runat="server" Text=""></asp:Label></div>
	
	  <div class="depdate">問題建置日期：<asp:Label ID="xPostDateText" runat="server" Text=""></asp:Label></div>
		
	  <!-- 頁面功能 Start -->
	  <ul class="Function2">
		  <li><asp:Label ID="TraceAddText" runat="server" Text=""></asp:Label></li>
		  <li><a href="javascript:print();" class="Print">友善列印</a></li>
		  <!--li><a href="#" class="Forward">轉寄好友</a></li-->
		  <li><asp:Label ID="BackText" runat="server" Text=""></asp:Label></li>
	  </ul>
	  <!-- 頁面功能 end -->
	
	  <!-- 內容區 Start -->
	
    <div id="cp">
    
        <!-- 內容區 Start -->
        <asp:Label ID="QuestionText" runat="server" Text=""></asp:Label>
        
  	    <!--駐站專家補充 /開始 -->
  	    <asp:Label ID="ProsText" runat="server" Text=""></asp:Label>
  	    
  	   <div class="btnCenter">
            <asp:Button ID="DeleteBtn" CssClass="btn2" runat="server" Text="刪除" visible="false" />            
	    </div>	
	    	    
	  </div>
    <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td>
   
</asp:Content>


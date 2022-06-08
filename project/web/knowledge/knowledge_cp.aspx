<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="knowledge_cp.aspx.vb" Inherits="knowledge_cp" title="農業知識入口網-農業知識家" ValidateRequest="false" %>

<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <!--Modify by Max 　農業字典-->
    <script src="../js/jtip.js" type="text/javascript"></script>
    <style type="text/css" media="all">
    @import "../css/global.css";
    .style1
    {
        height: 23px;
    }
</style>
  <script type="text/javascript">

    function btnDel(btn) {
        var key = btn.id.substring(btn.id.indexOf('_') + 1);
        var btn = document.getElementById("ctl00_ContentPlaceHolder1_btnDelProsText");
        if (confirm('確定要刪除回覆內容?')) {
            document.getElementById("ctl00_ContentPlaceHolder1_hidProsText").value = key;
            btn.click();
        }
    }
    function btnDelPic(btn) {
        var key = btn.id.substring(btn.id.indexOf('_') + 1);
        var btn = document.getElementById("ctl00_ContentPlaceHolder1_btnDelPic");
        if (confirm('確定要刪除圖片?')) {
            document.getElementById("ctl00_ContentPlaceHolder1_hidProsText").value = key;
            btn.click();
        }
    }
    

    function Question() {
        <%if String.IsNullOrEmpty(Session("memID")) then %>
            alert("請先登入會員");
        <%else %>
            var now = new Date();
            var domainName = document.domain;
            window.showModalDialog('http://' + domainName + '/mailbox.asp?Question=123&T=' + now.getSeconds(), self);             
        <%end if %>        
        return false;
    }
     

    function Discuss(id) {        
        <%if String.IsNullOrEmpty(Session("memID")) then %>
            alert("請先登入會員");
        <%else %>
            var now = new Date();
            var domainName = document.domain;
            window.showModalDialog('http://' + domainName + '/mailbox.asp?jDiscuss=&DiscussID=' + id + '&T=' + now.getSeconds(), self);                         
        <%end if %>        
        return false;
    }        
  </script>

  <uc1:LeftMenu ID="LeftMenu1" runat="server" />
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

    <div class="path">目前位置：<a href="/knowledge/knowledge.aspx">首頁</a><asp:Label ID="PathText" runat="server" Text=""></asp:Label></div>
	
	  <div class="depdate">問題建置日期：<asp:Label ID="xPostDateText" runat="server" Text=""></asp:Label></div>
		
	  <!-- 頁面功能 Start -->
	  <ul class="Function2">
		  <li><asp:Label ID="TraceAddText" runat="server" Text=""></asp:Label></li>
		  <!--added by Joey, 20090914, 新增問題反應 -->
		  <li><a href="#" class="Rword" onclick="Question()">系統問題</a></li> 		  
		  <asp:HiddenField id="ARTICLE_ID_Question" runat="server" />
		  <li><a href="#" onclick="window.open('knowledge_cpPrint.aspx<%=Server.UrlEncode(Request.Url.Query).Replace("%3f","?").Replace("%3d","=").Replace("%26","&") %>')" class="Print">友善列印</a></li>
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
	    	
	    <!--討論/開始 -->
	    <asp:Label ID="DiscussText" runat="server" Text=""></asp:Label>
	    
	    <!--相似問題/開始 -->
	    <asp:Label ID="SimilarProblemText" runat="server" Text=""></asp:Label>
	  </div>
	  <asp:Button id="btnDelProsText" runat="server" Text="Button" BackColor="White" BorderStyle="None" ForeColor="White" /><asp:HiddenField ID="hidProsText" runat="server" />
	  <asp:Button id="btnDelPic" runat="server" Text="Button" BackColor="White" BorderStyle="None" style="display: none;" />
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td>
  <td><%=innerTag%></td>
  <td><%=accountTag%></td>
</asp:Content>


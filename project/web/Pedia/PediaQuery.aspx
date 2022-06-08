<%@ Page Language="VB" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="false" CodeFile="PediaQuery.aspx.vb" Inherits="Pedia_PediaQuery" title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <a href="#" title="網頁內容資料" class="Accesskey" accesskey="C">:::</a>
  
  <div class="path">目前位置：<a href="/">首頁</a>&gt;<a href="/Pedia/PediaList.aspx">農業知識小百科</a></div>
	
	<ul class="Function2">
		<li><a href="javascript:history.back()" class="Back">回上一頁</a></li>
	</ul>
	
	<!-- 內容區 Start -->
	<div class="pantology">
    <div class="head"></div>
    <div class="body">
	  <h2>農業知識小百科－單元查詢</h2>
    
    <table class="type01" summary="結果條列式">
	  <tr>
	    <th><label for="content2">詞目</label></th>
	    <td align="left"><asp:TextBox ID="WordTitle" runat="server" CssClass="input" onclick="clearText(this)">請輸入關鍵字</asp:TextBox></td>
	  </tr>
	  <tr>
	    <th><label for="content">英文詞目</label></th>
	    <td align="left"><asp:TextBox ID="WordEngTitle" runat="server" CssClass="input" onclick="clearText(this)">請輸入關鍵字</asp:TextBox></td>
	  </tr>
	  <tr>
	    <th><label for="content">名詞釋義</label></th>
	    <td align="left"><asp:TextBox ID="WordBody" runat="server" CssClass="input" onclick="clearText(this)">請輸入關鍵字</asp:TextBox></td>
	  </tr>
	  <tr>
      <th><label for="select2">開放補充</label></th>
	    <td align="left">
        <asp:DropDownList ID="IsOpenDDL" runat="server">
          <asp:ListItem Value="" Text="請選擇" Selected="true" />
          <asp:ListItem Value="Y" Text="是" />
          <asp:ListItem Value="N" Text="否" />
        </asp:DropDownList>	    
      </td>
	  </tr>	
    </table>
    <div class="s01">
      <asp:Button ID="SubmitBtn" runat="server" Text="查詢" CssClass="button02" />
      <asp:Button ID="ResetBtn" runat="server" Text="重設條件" CssClass="button03" />      
    </div>        
    </div>
    <div class="foot"></div>
	</div>
  <script language="javascript" type="text/javascript">
  
    function clearText(item) {
      if( item.value == "請輸入關鍵字" )
      {
        item.value = "";
      }
    }
  </script>
  
</asp:Content>


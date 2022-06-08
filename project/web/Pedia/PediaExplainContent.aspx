<%@ Page Language="VB" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="false" CodeFile="PediaExplainContent.aspx.vb" Inherits="Pedia_PediaExplainContent" title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <a href="#" title="網頁內容資料" class="Accesskey" accesskey="C">:::</a>
  
  <div class="path">目前位置：<a href="/">首頁</a>&gt;<a href="/Pedia/PediaList.aspx">農業知識小百科</a></div>
	
	<ul class="Function2">
		<li><asp:HyperLink ID="AddLink" runat="server" CssClass="add" Text="我來補充" /></li>
    <li><asp:HyperLink ID="BackToContent" runat="server" CssClass="word" Text="回本詞目" /></li>
		<li><a href="#" class="Print">友善列印</a></li>		
	</ul>
	
	<div class="pantology">
    <div class="head"></div>
    <div class="body">
	    <h2>農業知識小百科－補充解釋</h2>
      <table class="type01" summary="結果條列式">
      <tr>
        <th width="15%">詞目</th>
        <td><asp:Label ID="WordTitle" runat="server" /></td>
      </tr>
      <tr>
        <th>英文詞目</th>
        <td><asp:Label ID="WordEngTitle" runat="server" /></td>
      </tr>
      <tr>
        <th>補充發表者</th>
        <td><asp:Label ID="MemberName" runat="server" /></td>
      </tr>
      <tr>
        <th>補充發佈日期</th>
        <td><asp:Label ID="ExplainDate" runat="server" /></td>
      </tr>
      <tr>
        <th>補充解釋</th>
        <td><asp:Label ID="ExplainBody" runat="server" /></td>
      </tr>
      </table>
      <div class="top"><a href="#" title="top">top</a></div>   
    </div>
    <div class="foot"></div>
	</div>
	<script language="javascript" type="text/javascript">
  function WordSearch(word)
  {    
		document.SearchForm.Keyword.value = word;			            
    document.SearchForm.action = "/kp.asp?xdURL=Search/SearchResultList.asp&mp=1";
    document.SearchForm.submit();
  }
</script>
</asp:Content>


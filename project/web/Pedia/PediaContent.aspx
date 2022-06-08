<%@ Page Language="VB" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="false" CodeFile="PediaContent.aspx.vb" Inherits="Pedia_PediaContent" title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
 <link rel="stylesheet" href="../css/pedia.css" type="text/css" />
  
    
  <a href="#" title="網頁內容資料" class="Accesskey" accesskey="C">:::</a>
  
  <div class="path">目前位置：<a href="/">首頁</a>&gt;<a href="/Pedia/PediaList.aspx">農業知識小百科</a></div>
	
	<ul class="Function2">  
		<li><a onclick="window.open('PediaContentPrint.aspx<%=Request.Url.Query %>')" href="#" class="Print">友善列印</a></li>
		<li><a href="javascript:history.back()" class="Back">回上一頁</a></li>
	</ul>
	
	<div class="pantology">
    <div class="head"></div>
    <div class="body">
	    <h2>農業知識小百科</h2>
      <table class="type01" summary="結果條列式">
      <tr>
        <th width="15%">詞目</th>
        <td><asp:Label ID="WordTitle" runat="server" /></td>
      </tr>
      <tr style='display:<%=haveEngTitle%>;'>
        <th>英文詞目</th>
        <td><asp:Label ID="WordEngTitle" runat="server" /></td>
      </tr>
      <tr style='display:none;'>
        <th>學名(中/英)</th>
        <td><asp:Label ID="WordFormalName" runat="server" /></td>
      </tr>
      <tr style='display:none;'>
        <th>俗名(中/英)</th>
        <td><asp:Label ID="WordLocalName" runat="server" /></td>
      </tr>
      <tr style='display:<%=haveWordBody%>;'>
        <th>名詞釋義</th>
        <td><asp:Label ID="WordBody" runat="server" /></td>
      </tr>
      <tr style='display:<%=haveWordKeyword%>;'>
        <th>相關詞</th>
        <td><asp:Label ID="WordKeyword" runat="server" /></td>
      </tr>
      </table>
      <hr />
	  <ul class="Function2">
		  <li><asp:HyperLink ID="CanAddLink" runat="server" CssClass="add" Visible="false">我來補充</asp:HyperLink></li>
		  <li><asp:HyperLink ID="ExplainListLink" runat="server" CssClass="add_more" Visible="false">看全部</asp:HyperLink></li>
	  </ul>	
    <h2>補充解釋：</h2>
    
      <asp:Label ID="ExplainText" runat="server" />
    <div class="top"><a href="#" title="top">top</a></div>   
    </div>
    <div class="foot"></div>
	</div>
	<!-- 內容區 End -->
<%=accountTag%>
<script language="javascript" type="text/javascript">
  function WordSearch(word)
  {    
		document.SearchForm.Keyword.value = word;			            
    document.SearchForm.action = "/kp.asp?xdURL=Search/SearchResultList.asp&mp=1";
    document.SearchForm.submit();
  }
</script>

</asp:Content>


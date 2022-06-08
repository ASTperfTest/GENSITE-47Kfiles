<%@ Page Language="VB" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="false" CodeFile="PediaExplainAdd.aspx.vb" Inherits="Pedia_PediaExplainAdd" title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <a href="#" title="網頁內容資料" class="Accesskey" accesskey="C">:::</a>
  
  <div class="path">目前位置：<a href="/">首頁</a>&gt;<a href="/Pedia/PediaList.aspx">農業知識小百科</a></div>
	
	<ul class="Function2">
		<li><a href="javascript:history.back()" class="Back">回上一頁</a></li>
	</ul>
	
	<div class="pantology">
    <div class="head"></div>
    <div class="body">
  	  <h2>農業知識小百科－我要補充</h2>
    
      <table class="type01" summary="結果條列式">
      <tr>
        <th>詞目</th>
        <td><asp:HyperLink ID="WordTitle" runat="server"></asp:HyperLink></td>
      </tr>
      <tr>
        <th>英文詞目</th>
        <td><asp:Label ID="WordEngTitle" runat="server" Text="Label"></asp:Label></td>
      </tr>
      <tr>
        <th><label for="add">我的補充</label></th>
        <td><asp:TextBox ID="ExplainContent" runat="server" CssClass="txt" Columns="60" Rows="10" TextMode="MultiLine" /></td>
      </tr>
      <tr>
        <th rowspan="2" nowrap="nowrap" scope="row">機器人辨識：</th>
        <td><asp:Image ID="MemberCaptChaImage" runat="server" ImageUrl="" Width="200" Height="50" /><input type="hidden" value="<%=guid %>" id="MemberGuid" name="MemberGuid" /> </td>
      </tr>
      <tr>
        <td><asp:TextBox ID="CaptChaText" Columns="20" CssClass="txt" runat="server" /><br />請輸入上方圖片中的數字</td>
      </tr>    
      </table>
      <div class="s01">
        <asp:Button ID="SubmitBtn" runat="server" Text="送出" CssClass="button01" />        
        <asp:Button ID="TempSaveBtn" runat="server" Text="暫存" CssClass="button02" visible="false" />
        <asp:Button ID="CancelBtn" runat="server" Text="取消" CssClass="button03" />                
      </div>        
    </div>
    <div class="foot"></div>
	</div>	

</asp:Content>


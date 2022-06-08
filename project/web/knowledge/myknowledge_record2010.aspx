<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="myknowledge_record2010.aspx.vb" Inherits="knowledge_myknowledge_record2010" title="Untitled Page" ValidateRequest="false" %>

<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <link href="include/style.css" rel="stylesheet" type="text/css" />
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

	  <!--path star (style B)-->
	  <div class="path">目前位置：<a href="/knowledge/knowledge.aspx">首頁</a>&gt;<a href=""#"">我的個人紀錄&nbsp;(2010KPI活動積分)</a></div>
	  <!--path end-->

	  <h3>我的個人紀錄&nbsp;(2010KPI活動積分)</h3>
	
	  <!-- 我的發問TAB LIST／開始 -->
	  <div id="Magazine">
		<div id="MagTabs"><asp:Label ID="TabText" runat="server" Text=""></asp:Label></div>	
		<div class="Event">

		  <!--瀏覽排序/開始 -->
		  <h5><asp:label runat="server" ID="memberIdText"></asp:label> 的知識檔案</h5>
		  <!--瀏覽排序/結束 -->

      <table width="100%" class="myset_layout">
      <tr>
        <td width="20%">
          <br/>
	        <asp:image runat="server" ID="memberPicPath"></asp:image>
	        <br/><br/>	        
          <asp:imagebutton runat="server" ID="memberChangePicBtn"></asp:imagebutton>  
			    <br/><br/><br/>	<br/>		
			<asp:label runat="server" ID="activityGrade"></asp:label>
		  
        </td>
        <td align="left">
		      <!--表格條列 /開始 -->
		      <asp:label runat="server" ID="tableContentText"></asp:label>
		      <!--表格條列 /結束 -->
		      <br />
		      <center></center>
		    </td>
      </tr>
      </table>
	  </div>
	  </div>
    <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td>	  

</asp:Content>


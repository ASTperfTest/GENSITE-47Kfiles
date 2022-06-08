<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="myknowledge_discuss.aspx.vb" Inherits="myknowledge_discuss" title="農業知識入口網-農業知識家" ValidateRequest="false" %>

<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

    <!--path star (style B)-->
	  <div class="path">目前位置：<a href="/knowledge/knowledge.aspx">首頁</a>&gt;<a href=""#"">我的討論</a></div>
	
  	<h3>我的討論</h3>
	
  	<!-- 頁面功能 Start -->
  	<ul class="Function2">
		  <li><asp:HyperLink ID="GoBackLink" runat="server" CssClass="Back">回上一頁</asp:HyperLink></li>
	  </ul>
	
	  <!--內容區 開始 -->
	  <div class="mycp">
		
		<h4>標題：<asp:Label ID="QuestionTitleText" runat="server" Text=""></asp:Label></h4>
		
		<!--depdate star (style C)-->
		<div class="depdate">問題建置日期：<asp:Label ID="QuestionPostDateText" runat="server" Text=""></asp:Label></div>		
		
		<!--depsort star (style C)-->
		<div class="depsort">問題建置分類：<asp:Label ID="QuestionCategoryText" runat="server" Text=""></asp:Label></div>		
		
	    <!--發問管理/開始 -->
		<div class="problem2"><asp:Label ID="QuestionBodyText" runat="server" Text=""></asp:Label></div>
	
		<h5>討論一覽：</h5>
	    <!--我的討論/開始 -->
	    <input type="hidden" name="Atype" value="" />
		<input type="hidden" name="DArticleId" value="" />
		<input type="hidden" name="ArticleContent" value="" />
		<input type="hidden" name="PostTime" value="" />
		<asp:Label ID="QuestionDiscussText" runat="server" Text=""></asp:Label>
		
		<div class="btnCenter">
      <asp:Button ID="BackToDiscussListBtn" CssClass="btn2" runat="server" Text="返回討論列表" />		    	    	      
	  </div>		
	  </div>
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td>
		
	<script language="javascript">
	
	    function FormSubmit( atype, value, postTime, name )
	    {	        
	        document.aspnetForm.Atype.value = atype;
	        document.aspnetForm.DArticleId.value = value;
	        document.aspnetForm.PostTime.value = postTime;
	        document.aspnetForm.ArticleContent.value = document.getElementById(name).value;
	        document.aspnetForm.method = "post";	        
	        document.aspnetForm.submit();
	    }
	</script>
</asp:Content>


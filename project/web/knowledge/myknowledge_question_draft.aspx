<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="myknowledge_question_draft.aspx.vb" Inherits="myknowledge_question_draft" title="農業知識入口網-農業知識家" ValidateRequest="false" %>

<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

    <!--path star (style B)-->
	  <div class="path">目前位置：<a href="/knowledge/knowledge.aspx">首頁</a>&gt;<a href=""#"">我的發問</a></div>
	  <!--path end-->
	
	  <h3>我的發問</h3>
	
	  <!-- 頁面功能 Start -->
	  <ul class="Function2">
		  <li><asp:HyperLink ID="GoBackLink" runat="server" CssClass="Back">回上一頁</asp:HyperLink></li>
	  </ul>
	
	  <!--內容區 開始 -->
	  <div class="mycp">
		
		<h4>標題：<asp:Label ID="QuestionTitleText" runat="server" Text=""></asp:Label></h4>
		
		<!--depdate star (style C)-->
		<div class="depdate">問題建置日期：<asp:Label ID="QuestionPostDateText" runat="server" Text=""></asp:Label></div>
		<!--depdate end-->
		
		<!--depsort star (style C)-->
		<div class="depsort">問題建置分類：<asp:Label ID="QuestionCategoryText" runat="server" Text=""></asp:Label></div>
		<!--depsort end-->
		
		<table class="table1" border="0" cellspacing="0" cellpadding="0" summary="排版用表格">
		  <tr>
			<td>瀏覽：<asp:Label ID="QuestionBrowseCountText" CssClass="number2" runat="server" Text=""></asp:Label>次</td>
			<td>討論：<asp:Label ID="QuestionDiscussCountText" CssClass="number2" runat="server" Text=""></asp:Label>次</td>
			<td>追蹤：<asp:Label ID="QuestionTraceCountText" CssClass="number2" runat="server" Text=""></asp:Label>次</td>
		  </tr>
		  <tr>
			<td>意見：<asp:Label ID="QuestionCommandCountText" CssClass="number2" runat="server" Text=""></asp:Label>次</td>
			<td colspan="2">平均評價：
			    <asp:Label ID="AverageGradeCountText" runat="server" Text=""></asp:Label>
			</td>
		  </tr>
		</table>

	    <!--發問管理/開始 -->
		<h5>發問管理：</h5>
	
	    <!--我要補充/開始 -->
		<div class="add">		  
			<label for="textarea">我的發問：</label>
			<asp:TextBox ID="QuestionAdditionalContentText" TextMode="multiLine" Rows="10" runat="server"></asp:TextBox>			
			<div class="btnCenter">
                <asp:Button ID="QuestionConfirmBtn" CssClass="btn2" runat="server" Text="確定發佈" />
			    <asp:Button ID="QuestionSaveBtn" CssClass="btn2" runat="server" Text="儲存草稿" />			    
			    <asp:Button ID="QuestionDelBtn" CssClass="btn2" runat="server" Text="刪除發問" OnClientClick="ConfirmDelete();" />
			</div>	      
	    </div>			
	  </div>
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td>	 
	  
	<script language="javascript" type="text/javascript">
	
	    function ConfirmDelete() {
	    
	        if(window.confirm("請問您確定要刪除此篇文章?") != true) {
	            event.returnValue = false;
	        }
	    }
	</script>


</asp:Content>


<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="knowledge_discuss_cp.aspx.vb" Inherits="knowledge_discuss_cp" title="農業知識入口網-農業知識家" ValidateRequest="false" Async="true" %>

<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
  <!--Modify by Max 　調整字級-->
    <script type="text/javascript" src="/subject/js/dozoom.js"></script>
  <!--End-->  
  <!--Modify by Max 　農業字典-->
  <script src="../js/jtip.js" type="text/javascript"></script>
  <style type="text/css" media="all">
  @import "../css/global.css";
  .style1
  {
      height: 23px;
  }
  </style>
  <!--End--> 
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
	<!--Modify by Max 　調整字級-->
	  <DIV align="right">			              
           調整字級：
                <A onclick="changeFontSize('article','idx1');">
                    <IMG id="idx1" alt="" src="../subject/images/fontsize_1_off.gif" align="absMiddle" border="0" name="idx1" style="padding-left:0px;padding-right:0px;padding-bottom:2px;" />
                </A>
                <A onclick="changeFontSize('article','idx2');">
                    <IMG id="idx2" alt="" src="../subject/images/fontsize_2_off.gif" align="absMiddle" border="0" name="idx2" style="padding-left:0px;padding-right:0px;padding-bottom:2px;" />
                </A>
                <A onclick="changeFontSize('article','idx3');">
                    <IMG id="idx3" alt="" src="../subject/images/fontsize_3_off.gif" align="absMiddle" border="0" name="idx3" style="padding-left:0px;padding-right:0px;padding-bottom:2px;" />
                </A>
                <A onclick="changeFontSize('article','idx4');">
                    <IMG id="idx4" alt="" src="../subject/images/fontsize_4_off.gif" align="absMiddle" border="0" name="idx4" style="padding-left:0px;padding-right:0px;padding-bottom:2px;" />
                </A>				
      </DIV>
      <!--End-->
	  <!-- 頁面功能 end -->
	
	  <!-- 內容區 Start -->
	  <div id="cp">
	  <span id="article" name="article" class="idx1">
	    <asp:Label ID="QuestionText" runat="server" Text=""></asp:Label>
	
	    <!--討論/開始 -->
	    <asp:Label ID="ArticleText" runat="server" Text=""></asp:Label>
	
	    <!--評價顯示/開始 -->		
	    <div id="stararea" runat="server">
	    <div id="ratingAndStatsDiv">
		    您對本篇補充的評價：
		    <div id="ratingdiv">
		      (有待加強)
		      <a href="#rating"></a>
		      <asp:HyperLink ID="Star1ImgLink" runat="server" NavigateUrl="#rating" CssClass="rating" onmouseover="checkstar('1')">
                  <asp:Image ID="Star1Img" CssClass="rating" ImageUrl="images/icn_star_empty_19x20.gif" runat="server" />
		      </asp:HyperLink>
              <asp:HyperLink ID="Star2ImgLink" runat="server" NavigateUrl="#rating" CssClass="rating" onmouseover="checkstar('2')">
                  <asp:Image ID="Star2Img" CssClass="rating" ImageUrl="images/icn_star_empty_19x20.gif" runat="server" />
		      </asp:HyperLink>
		      <asp:HyperLink ID="Star3ImgLink" runat="server" NavigateUrl="#rating" CssClass="rating" onmouseover="checkstar('3')">
                  <asp:Image ID="Star3Img" CssClass="rating" ImageUrl="images/icn_star_empty_19x20.gif" runat="server" />
		      </asp:HyperLink>
		      <asp:HyperLink ID="Star4ImgLink" runat="server" NavigateUrl="#rating" CssClass="rating" onmouseover="checkstar('4')">
                  <asp:Image ID="Star4Img" CssClass="rating" ImageUrl="images/icn_star_empty_19x20.gif" runat="server" />
		      </asp:HyperLink>
		      <asp:HyperLink ID="Star5ImgLink" runat="server" NavigateUrl="#rating" CssClass="rating" onmouseover="checkstar('5')">
                  <asp:Image ID="Star5Img" CssClass="rating" ImageUrl="images/icn_star_empty_19x20.gif" runat="server" />
		      </asp:HyperLink>		                    
		      (非常有價值)  
		    </div>
	    </div>        
	    </div>
	    <!--評價顯示/結束 -->
	    <br />
	    <!--發表意見/開始 -->
 	    <label for="textarea"><asp:Label ID="PostCommentLabel" runat="server" Text=""></asp:Label></label>
	    <br />
	    <asp:TextBox ID="PostCommentTBox" TextMode="MultiLine" Rows="10" Columns="60" runat="server" onfocus="checkfocus(this);" Wrap="true">請在這裡輸入您對本篇討論的意見(限150字)</asp:TextBox>
	    <asp:TextBox ID="StarValue" Width="0" Height="0" BorderWidth="0" Text="0" runat="server"></asp:TextBox>	
	    <br /><br />
        <label for="textfield"><asp:Label ID="CaptChaText1" runat="server" Text="">機器人辨識</asp:Label></label>
        <asp:Image ID="MemberCaptChaImage" runat="server" ImageUrl="" />
        <input type="hidden" value="<%=guid %>" id="MemberGuid" name="MemberGuid" />  
        <br />
        <label for="textfield"><asp:Label ID="CaptChaText2" runat="server" Text="">請輸入上方圖片的數字</asp:Label></label>
    	<asp:TextBox ID="MemberCaptChaTBox" Columns="20" CssClass="txt" runat="server" />
        <br /><br />	
        <div>
            <asp:Button ID="ConfirmBtn" runat="server" Text="確定" CssClass="btn2" OnClientClick="CheckSubmit()" />
            <asp:Button ID="CancelBtn" runat="server" Text="取消" CssClass="btn2" />
	    </div>
	    <!--發表意見/結束 -->    	
	    <!--討論/開始 -->
	    <asp:Label ID="DiscussText" runat="server" Text=""></asp:Label>
	    <!--討論/結束 -->
	  </div>
	  </span>
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
	</td>
	
	<script language="javascript" type="text/javascript">
	
	    function checkstar(value) {
	        if(value == "5") {
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star1Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star2Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star3Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star4Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star5Img.src = "images/icn_star_full_19x20.gif";
	        }
	        if(value == "4") {
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star1Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star2Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star3Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star4Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star5Img.src = "images/icn_star_empty_19x20.gif";
	        }
	        if(value == "3") {
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star1Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star2Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star3Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star4Img.src = "images/icn_star_empty_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star5Img.src = "images/icn_star_empty_19x20.gif";
	        }
	        if(value == "2") {
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star1Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star2Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star3Img.src = "images/icn_star_empty_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star4Img.src = "images/icn_star_empty_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star5Img.src = "images/icn_star_empty_19x20.gif";
	        }
	        if(value == "1") {
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star1Img.src = "images/icn_star_full_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star2Img.src = "images/icn_star_empty_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star3Img.src = "images/icn_star_empty_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star4Img.src = "images/icn_star_empty_19x20.gif";
	            document.aspnetForm.ctl00_ContentPlaceHolder1_Star5Img.src = "images/icn_star_empty_19x20.gif";
	        }	        
	        document.aspnetForm.ctl00_ContentPlaceHolder1_StarValue.value = value;
	    }
	
	</script>
		<script language="javascript" type="text/javascript">
	
	    function checkfocus(item) {
	    
	        if(item.value == "請在這裡輸入您對本篇討論的意見(限150字)"){
	            item.value = "";
	        }
	    
	    }	    	 
	    function CheckSubmit() {
	      if( document.getElementById("ctl00_ContentPlaceHolder1_PostCommentTBox").value == "請在這裡輸入您對本篇討論的意見(限150字)" )
	      {
	        alert("請輸入您的意見");
	        window.event.returnValue = false;
	      }
	      else if( document.getElementById("ctl00_ContentPlaceHolder1_PostCommentTBox").value == "" )
	      {
	        alert("請輸入您的意見");
	        window.event.returnValue = false;
	      }
	      else
	      {
	        window.event.returnValue = true;
	      }
	    }
	</script>
    <td><%=innerTag%></td>
</asp:Content>


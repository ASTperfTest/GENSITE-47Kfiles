<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="knowledge_professional_add.aspx.vb" Inherits="knowledge_professional_add" title="農業知識入口網-農業知識家" ValidateRequest="false" %>

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
        
        <!--發表意見/開始 -->
        <label for="textfield"><asp:Label ID="PostCommentLabel" runat="server" Text=""></asp:Label></label>
        <div id="discusspage" runat="server">
            <br />
		    <label for="textfield">專家補充：</label>
		    <br />
			<asp:TextBox ID="DiscussContent" TextMode="MultiLine" Rows="9" width="100%" runat="server"  onfocus="checkfocus(this);">請在這裡輸入您對本問題的補充</asp:TextBox>	
			<br />
            <asp:Label ID="textfield" runat="server" Text="上傳圖片："></asp:Label>
            <asp:Button ID="uploadBtn" CssClass="btn2" runat="server" Text="上傳圖片" OnClientClick=
                "window.showModalDialog('UploadImageDialog.aspx',self,'resizable:no;scrollbars:Yes;dialogWidth:900px;dialogHeight:600px;center:Yes');" />
            <asp:Button ID="deleteBtn" CssClass="btn2" runat="server" Text="刪除圖片"/>
			<br/>
            <asp:Image ID="previewImg" runat="server" Visible="False" />
            <br/>
            <asp:Label ID="lblImgHint" runat="server" Text=""></asp:Label>
            <div id="divPreviewImg" runat="server" style="width:450px;"></div>
            <asp:HiddenField ID="hidFileName" runat="server" />
            <asp:HiddenField ID="hidFileContent" runat="server" />
            <br />
            <label for="textfield">機器人辨識</label>
            <asp:Image ID="MemberCaptChaImage" runat="server" ImageUrl="" />
            <input type="hidden" value="<%=guid %>" id="MemberGuid" name="MemberGuid" /> 
            <br />
            <label for="textfield"><asp:Label ID="CaptChaText2" runat="server" Text="">請輸入上方圖片的數字</asp:Label></label>
    	    <asp:TextBox ID="MemberCaptChaTBox" Columns="20" CssClass="txt" runat="server" />
            <br /><br />	    
		    <!--p>請注意!! 討論一經「確定」即發表，無法進行刪除，請確實檢查您欲發表的內容，並對自己的言論負責。</p-->
		    <div>
                <asp:Button ID="SubmitBtn" CssClass="btn2" runat="server" Text="確定" />                
                <asp:Button ID="ResetBtn" CssClass="btn2" runat="server" Text="取消" />			
			</div>
	        <!--發表意見/結束 -->
	    </div>  	    
	  </div>
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td>
	  
	<script language="javascript" type="text/javascript">
	
	    function checkfocus(item) {
	    
	        if(item.value == "請在這裡輸入您對本問題的補充"){
	            item.value = "";
	        }

	    }
	    
	    function GetReturnValue(str, content) {
	        document.getElementById('ctl00_ContentPlaceHolder1_hidFileName').value += str + "^";
	        document.getElementById('ctl00_ContentPlaceHolder1_hidFileContent').value += content + "^";
	    }  
	</script>
</asp:Content>


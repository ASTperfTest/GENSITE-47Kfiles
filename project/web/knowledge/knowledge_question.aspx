<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="knowledge_question.aspx.vb" Inherits="knowledge_question" title="農業知識入口網-農業知識家" ValidateRequest="false" %>

<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <uc1:LeftMenu ID="LeftMenu1" runat="server" />
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

    <div class="path">目前位置：<a href="/knowledge/knowledge.aspx">首頁</a>&gt;<a href="/knowledge/knowledge_question.aspx">我要發問</a></div>
	
	  <h3>我要發問</h3>
	
	  <!-- 頁面功能 Start -->
	  <ul class="Function2">
		  <li><a href="javascript:history.back();" class="Back">回上一頁</a></li>
		</ul>
	  <!-- 頁面功能 end -->
	  
	  <!-- 內容區 Start -->
	
    <div class="mycp">
    
    <asp:label runat="server" id="RelationUnit" text=""></asp:label>
                 
    <!--我要補充/開始 -->
		<div class="add" id="add" runat="server">
				   
			<label for="textfield">問題標題：</label>			
      <asp:TextBox ID="QuestionTitle" CssClass="addtextfield" runat="server" onfocus="checkfocustitle(this);">請輸入清楚明白的問題標題，限 30 個字以內</asp:TextBox>           			
			
			<label for="textarea">問題內容：</label>
			<asp:TextBox ID="QuestionContent" TextMode="MultiLine" runat="server"  onfocus="checkfocuscontent(this);"  Wrap="true">請詳細說明問題的狀況</asp:TextBox>
								
			<label for="textfield">問題Tag：</label>
			<asp:TextBox ID="Tag1" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag2" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag3" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag4" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag5" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag6" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag7" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag8" runat="server"></asp:TextBox>			
			<br /><br /><br />
            <asp:Label ID="textfield" runat="server" Text="上傳圖片："></asp:Label>
            <asp:Button ID="uploadBtn" CssClass="btn2" runat="server" Text="上傳圖片" OnClientClick=
                "window.showModalDialog('UploadImageDialog.aspx',self,'resizable:no;scrollbars:Yes;dialogWidth:900px;dialogHeight:600px;center:Yes'); " />
            <asp:Button ID="deleteBtn" CssClass="btn2" runat="server" Text="刪除圖片" Visible="False" />
			<br/>
            <asp:Image ID="previewImg" runat="server" Visible="False" />
            <br/>
            <asp:Label ID="lblImgHint" runat="server" Text=""></asp:Label>
            <div id="divPreviewImg" runat="server" style="width:450px"></div>
			<!--<div runat="server" id="uploadPicDiv">
			  <label for="textfield">上傳圖片：</label>
				<div class="knowledgepic">
				  <ol>				   
                    <asp:panel runat="server" id="filePanel"> </asp:panel>
				  </ol>
		          <p class="importanttext">每張圖片檔案大小限制需小於200k，檔案格式限定.jpg或.gif。輸入圖片說明文字可幫助其他人更容易了解。<br/>
			          圖片上傳後，需先經過網站管理者審核，審核通過者，方會在前台發問或討論頁面呈現。</p>
			      <br />				  
				</div>
			</div>-->
            <asp:HiddenField ID="hidFileName" runat="server" />
            <asp:HiddenField ID="hidFileContent" runat="server" />
			<br/><br/><br />
			問題分類：
      <asp:DropDownList ID="QuestionTypeDDL" runat="server">                
        <asp:ListItem Value="A" Text="農" />
        <asp:ListItem Value="B" Text="林" />
        <asp:ListItem Value="C" Text="漁" />
        <asp:ListItem Value="D" Text="牧" />
        <asp:ListItem Value="E" Text="其他" />
        
      </asp:DropDownList>
      <br /><br />
	  <asp:panel id="pnlGarden" runat="server" algin="left" Visible="False" width="100%">
		園藝教室分類：
		<asp:DropDownList ID="GardenTypeDDL" runat="server">                
		</asp:DropDownList>
	  <br /><br />
	  </asp:panel>
      <label for="textfield">機器人辨識</label>
      <asp:Image ID="MemberCaptChaImage" runat="server" ImageUrl="" />
      <input type="hidden" value="<%=guid %>" id="MemberGuid" name="MemberGuid" />  
      <label for="textfield">請輸入上方圖片的數字</label>
    	<asp:TextBox ID="MemberCaptChaTBox" Columns="20" CssClass="txt" runat="server" />
      <br /><br />
			<p>請注意!! 問題一經「確定」發表，當有人對您的問題進行討論後即無法進行刪除，請確實檢查您欲發表的內容，並對自己的言論負責。</p>
			<div class="btnCenter">
        <asp:Button ID="SubmitBtn" CssClass="btn2" runat="server" Text="確定" />
        <asp:Button ID="TempSubmitBtn" CssClass="btn2" runat="server" Text="暫存" />
        <asp:Button ID="ResetBtn" CssClass="btn2" runat="server" Text="取消" />			
			</div>	 			
	  </div>
	  <!--我要補充/結束 -->	 	 	  	    
	  </div>
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td>
  
	<script language="javascript" type="text/javascript">	    
	  function checkfocustitle(item) {
	    if(item.value == "請輸入清楚明白的問題標題，限 30 個字以內"){
	      item.value = "";
	    }	    
	  }	  
	  function checkfocuscontent(item) {
	    if(item.value == "請詳細說明問題的狀況"){
	      item.value = "";
	    }	    
	  }
	
	  var sizeLimit = 200000; //sizeLimit單位:byte
		var globalflag = true;
		var picid = "";
		var picvalue = "";
				
	  function checkFileSize(nid) {
	    picid = "ctl00_ContentPlaceHolder1_picFile_" + nid;
	    if( document.getElementById(picid).value != "" ) {
	      picvalue = document.getElementById(picid).value;
	      var re = /(\.jpg|\.gif)$/i;             
        if (!re.test(picvalue)) {
          alert('只允許上傳JPG或GIF影像檔');    
          document.getElementById(picid).outerHTML += "";          
        } 
        else {
	        var img = new Image(); 
          img.sizeLimit = sizeLimit; 
          img.src = document.getElementById(picid).value;  
          img.onload = showImageDimensions;     
        }
	    }
	  }
	  function showImageDimensions() 
    {     
      if (this.fileSize > sizeLimit) {         
        alert('您所選擇的檔案大小為 ' + (this.fileSize/1000) +' kb，\n超過了上傳上限 ' + (sizeLimit/1000) + ' kb！\n不允許上傳！');         
        document.getElementById(picid).outerHTML += "";         
      }                  
    } 
	  function checkFile(count)
	  {		      
	    for( var i = 0; i < count; i++ ) { 
        var id = "ctl00_ContentPlaceHolder1_picFile_" + i;
        var FileId = "ctl00_ContentPlaceHolder1_picDesc_" + i;
        if( document.getElementById(id).value.length != 0 ) {
          if(document.getElementById(FileId).value == "請輸入圖片說明" || document.getElementById(FileId).value.length == 0 ) {
            alert('上傳圖片檔案 : ' + (i+1) + ", 請輸入檔案敘述");
            globalflag = false;
          }         
          else if(document.getElementById(id).value.length == 0 && document.getElementById(FileId).value.length != 0 && document.getElementById(FileId).value != "請輸入圖片說明" ) {
            alert('請上傳圖片檔案 :' + (i+1) );
            globalflag = false;
          } 
          else {
            globalflag = true;
          }                                               
        }
      }  
	    return globalflag;
	}

	function GetReturnValue(str, content) {
	    document.getElementById('ctl00_ContentPlaceHolder1_hidFileName').value += str + "^";
	    document.getElementById('ctl00_ContentPlaceHolder1_hidFileContent').value += content + "^";
	}  
	</script> 

</asp:Content>


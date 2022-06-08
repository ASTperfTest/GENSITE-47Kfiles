<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="knowledge_question.aspx.vb" Inherits="knowledge_question" title="�A�~���ѤJ�f��-�A�~���Ѯa" ValidateRequest="false" %>

<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <uc1:LeftMenu ID="LeftMenu1" runat="server" />
        
  <!--�������e��--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

    <div class="path">�ثe��m�G<a href="/knowledge/knowledge.aspx">����</a>&gt;<a href="/knowledge/knowledge_question.aspx">�ڭn�o��</a></div>
	
	  <h3>�ڭn�o��</h3>
	
	  <!-- �����\�� Start -->
	  <ul class="Function2">
		  <li><a href="javascript:history.back();" class="Back">�^�W�@��</a></li>
		</ul>
	  <!-- �����\�� end -->
	  
	  <!-- ���e�� Start -->
	
    <div class="mycp">
    
    <asp:label runat="server" id="RelationUnit" text=""></asp:label>
                 
    <!--�ڭn�ɥR/�}�l -->
		<div class="add" id="add" runat="server">
				   
			<label for="textfield">���D���D�G</label>			
      <asp:TextBox ID="QuestionTitle" CssClass="addtextfield" runat="server" onfocus="checkfocustitle(this);">�п�J�M�����ժ����D���D�A�� 30 �Ӧr�H��</asp:TextBox>           			
			
			<label for="textarea">���D���e�G</label>
			<asp:TextBox ID="QuestionContent" TextMode="MultiLine" runat="server"  onfocus="checkfocuscontent(this);"  Wrap="true">�иԲӻ������D�����p</asp:TextBox>
								
			<label for="textfield">���DTag�G</label>
			<asp:TextBox ID="Tag1" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag2" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag3" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag4" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag5" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag6" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag7" runat="server"></asp:TextBox>
			<asp:TextBox ID="Tag8" runat="server"></asp:TextBox>			
			<br /><br /><br />
            <asp:Label ID="textfield" runat="server" Text="�W�ǹϤ��G"></asp:Label>
            <asp:Button ID="uploadBtn" CssClass="btn2" runat="server" Text="�W�ǹϤ�" OnClientClick=
                "window.showModalDialog('UploadImageDialog.aspx',self,'resizable:no;scrollbars:Yes;dialogWidth:900px;dialogHeight:600px;center:Yes'); " />
            <asp:Button ID="deleteBtn" CssClass="btn2" runat="server" Text="�R���Ϥ�" Visible="False" />
			<br/>
            <asp:Image ID="previewImg" runat="server" Visible="False" />
            <br/>
            <asp:Label ID="lblImgHint" runat="server" Text=""></asp:Label>
            <div id="divPreviewImg" runat="server" style="width:450px"></div>
			<!--<div runat="server" id="uploadPicDiv">
			  <label for="textfield">�W�ǹϤ��G</label>
				<div class="knowledgepic">
				  <ol>				   
                    <asp:panel runat="server" id="filePanel"> </asp:panel>
				  </ol>
		          <p class="importanttext">�C�i�Ϥ��ɮפj�p����ݤp��200k�A�ɮ׮榡���w.jpg��.gif�C��J�Ϥ�������r�i���U��L�H��e���F�ѡC<br/>
			          �Ϥ��W�ǫ�A�ݥ��g�L�����޲z�̼f�֡A�f�ֳq�L�̡A��|�b�e�x�o�ݩΰQ�׭����e�{�C</p>
			      <br />				  
				</div>
			</div>-->
            <asp:HiddenField ID="hidFileName" runat="server" />
            <asp:HiddenField ID="hidFileContent" runat="server" />
			<br/><br/><br />
			���D�����G
      <asp:DropDownList ID="QuestionTypeDDL" runat="server">                
        <asp:ListItem Value="A" Text="�A" />
        <asp:ListItem Value="B" Text="�L" />
        <asp:ListItem Value="C" Text="��" />
        <asp:ListItem Value="D" Text="��" />
        <asp:ListItem Value="E" Text="��L" />
        
      </asp:DropDownList>
      <br /><br />
	  <asp:panel id="pnlGarden" runat="server" algin="left" Visible="False" width="100%">
		�����ЫǤ����G
		<asp:DropDownList ID="GardenTypeDDL" runat="server">                
		</asp:DropDownList>
	  <br /><br />
	  </asp:panel>
      <label for="textfield">�����H����</label>
      <asp:Image ID="MemberCaptChaImage" runat="server" ImageUrl="" />
      <input type="hidden" value="<%=guid %>" id="MemberGuid" name="MemberGuid" />  
      <label for="textfield">�п�J�W��Ϥ����Ʀr</label>
    	<asp:TextBox ID="MemberCaptChaTBox" Columns="20" CssClass="txt" runat="server" />
      <br /><br />
			<p>�Ъ`�N!! ���D�@�g�u�T�w�v�o��A���H��z�����D�i��Q�׫�Y�L�k�i��R���A�нT���ˬd�z���o�����e�A�ù�ۤv�����׭t�d�C</p>
			<div class="btnCenter">
        <asp:Button ID="SubmitBtn" CssClass="btn2" runat="server" Text="�T�w" />
        <asp:Button ID="TempSubmitBtn" CssClass="btn2" runat="server" Text="�Ȧs" />
        <asp:Button ID="ResetBtn" CssClass="btn2" runat="server" Text="����" />			
			</div>	 			
	  </div>
	  <!--�ڭn�ɥR/���� -->	 	 	  	    
	  </div>
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="�^�쭶���̤W��" /></a></div><!--20080829 chiung -->
  </td>
  
	<script language="javascript" type="text/javascript">	    
	  function checkfocustitle(item) {
	    if(item.value == "�п�J�M�����ժ����D���D�A�� 30 �Ӧr�H��"){
	      item.value = "";
	    }	    
	  }	  
	  function checkfocuscontent(item) {
	    if(item.value == "�иԲӻ������D�����p"){
	      item.value = "";
	    }	    
	  }
	
	  var sizeLimit = 200000; //sizeLimit���:byte
		var globalflag = true;
		var picid = "";
		var picvalue = "";
				
	  function checkFileSize(nid) {
	    picid = "ctl00_ContentPlaceHolder1_picFile_" + nid;
	    if( document.getElementById(picid).value != "" ) {
	      picvalue = document.getElementById(picid).value;
	      var re = /(\.jpg|\.gif)$/i;             
        if (!re.test(picvalue)) {
          alert('�u���\�W��JPG��GIF�v����');    
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
        alert('�z�ҿ�ܪ��ɮפj�p�� ' + (this.fileSize/1000) +' kb�A\n�W�L�F�W�ǤW�� ' + (sizeLimit/1000) + ' kb�I\n�����\�W�ǡI');         
        document.getElementById(picid).outerHTML += "";         
      }                  
    } 
	  function checkFile(count)
	  {		      
	    for( var i = 0; i < count; i++ ) { 
        var id = "ctl00_ContentPlaceHolder1_picFile_" + i;
        var FileId = "ctl00_ContentPlaceHolder1_picDesc_" + i;
        if( document.getElementById(id).value.length != 0 ) {
          if(document.getElementById(FileId).value == "�п�J�Ϥ�����" || document.getElementById(FileId).value.length == 0 ) {
            alert('�W�ǹϤ��ɮ� : ' + (i+1) + ", �п�J�ɮױԭz");
            globalflag = false;
          }         
          else if(document.getElementById(id).value.length == 0 && document.getElementById(FileId).value.length != 0 && document.getElementById(FileId).value != "�п�J�Ϥ�����" ) {
            alert('�ФW�ǹϤ��ɮ� :' + (i+1) );
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


<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="knowledge_professional_add.aspx.vb" Inherits="knowledge_professional_add" title="�A�~���ѤJ�f��-�A�~���Ѯa" ValidateRequest="false" %>

<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <script src="http://www.google-analytics.com/urchin.js" type="text/javascript"></script>

  <uc1:LeftMenu ID="LeftMenu1" runat="server" />
        
  <!--�������e��--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

    <div class="path">�ثe��m�G<a href="/knowledge/knowledge.aspx">����</a><asp:Label ID="PathText" runat="server" Text=""></asp:Label></div>
	
	  <div class="depdate">���D�ظm����G<asp:Label ID="xPostDateText" runat="server" Text=""></asp:Label></div>
		
	  <!-- �����\�� Start -->
	  <ul class="Function2">
  		<li><asp:Label ID="TraceAddText" runat="server" Text=""></asp:Label></li>
		  <li><a href="javascript:print();" class="Print">�͵��C�L</a></li>
		  <!--li><a href="#" class="Forward">��H�n��</a></li-->
		  <li><asp:Label ID="BackText" runat="server" Text=""></asp:Label></li>
	  </ul>
	  <!-- �����\�� end -->
	
	  <!-- ���e�� Start -->	
    <div id="cp">
    
        <!-- ���e�� Start -->
        <asp:Label ID="QuestionText" runat="server" Text=""></asp:Label>
        
        <!--�o��N��/�}�l -->
        <label for="textfield"><asp:Label ID="PostCommentLabel" runat="server" Text=""></asp:Label></label>
        <div id="discusspage" runat="server">
            <br />
		    <label for="textfield">�M�a�ɥR�G</label>
		    <br />
			<asp:TextBox ID="DiscussContent" TextMode="MultiLine" Rows="9" width="100%" runat="server"  onfocus="checkfocus(this);">�Цb�o�̿�J�z�糧���D���ɥR</asp:TextBox>	
			<br />
            <asp:Label ID="textfield" runat="server" Text="�W�ǹϤ��G"></asp:Label>
            <asp:Button ID="uploadBtn" CssClass="btn2" runat="server" Text="�W�ǹϤ�" OnClientClick=
                "window.showModalDialog('UploadImageDialog.aspx',self,'resizable:no;scrollbars:Yes;dialogWidth:900px;dialogHeight:600px;center:Yes');" />
            <asp:Button ID="deleteBtn" CssClass="btn2" runat="server" Text="�R���Ϥ�"/>
			<br/>
            <asp:Image ID="previewImg" runat="server" Visible="False" />
            <br/>
            <asp:Label ID="lblImgHint" runat="server" Text=""></asp:Label>
            <div id="divPreviewImg" runat="server" style="width:450px;"></div>
            <asp:HiddenField ID="hidFileName" runat="server" />
            <asp:HiddenField ID="hidFileContent" runat="server" />
            <br />
            <label for="textfield">�����H����</label>
            <asp:Image ID="MemberCaptChaImage" runat="server" ImageUrl="" />
            <input type="hidden" value="<%=guid %>" id="MemberGuid" name="MemberGuid" /> 
            <br />
            <label for="textfield"><asp:Label ID="CaptChaText2" runat="server" Text="">�п�J�W��Ϥ����Ʀr</asp:Label></label>
    	    <asp:TextBox ID="MemberCaptChaTBox" Columns="20" CssClass="txt" runat="server" />
            <br /><br />	    
		    <!--p>�Ъ`�N!! �Q�פ@�g�u�T�w�v�Y�o��A�L�k�i��R���A�нT���ˬd�z���o�����e�A�ù�ۤv�����׭t�d�C</p-->
		    <div>
                <asp:Button ID="SubmitBtn" CssClass="btn2" runat="server" Text="�T�w" />                
                <asp:Button ID="ResetBtn" CssClass="btn2" runat="server" Text="����" />			
			</div>
	        <!--�o��N��/���� -->
	    </div>  	    
	  </div>
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="�^�쭶���̤W��" /></a></div><!--20080829 chiung -->
  </td>
	  
	<script language="javascript" type="text/javascript">
	
	    function checkfocus(item) {
	    
	        if(item.value == "�Цb�o�̿�J�z�糧���D���ɥR"){
	            item.value = "";
	        }

	    }
	    
	    function GetReturnValue(str, content) {
	        document.getElementById('ctl00_ContentPlaceHolder1_hidFileName').value += str + "^";
	        document.getElementById('ctl00_ContentPlaceHolder1_hidFileContent').value += content + "^";
	    }  
	</script>
</asp:Content>


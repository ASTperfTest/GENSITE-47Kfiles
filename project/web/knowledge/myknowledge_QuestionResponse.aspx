<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="myknowledge_QuestionResponse.aspx.vb" Inherits="myknowledge_QuestionResponse" title="Untitled Page" ValidateRequest="false" %>

<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
        
  <!--�������e��--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />
	
	  <!--path star (style B)-->
	  <div class="path">�ثe��m�G<a href="/knowledge/knowledge.aspx">����</a>&gt;<a href=""#"">�ڤ��������D</a></div>
	  <!--path end-->

	  <h3>�ڤ��������D</h3>
	
	  <!-- �ڪ��o��TAB LIST���}�l -->
	  <div id="Magazine">
		<div id="MagTabs"><asp:Label ID="TabText" runat="server" Text=""></asp:Label></div>	
		<div class="Event">
      <div><!--
			  <label>��ƿz��G
			    <a href="/knowledge/myknowledge_pedia.aspx?type=A">�ڪ����˵��J</a>�U<a href="/knowledge/myknowledge_pedia.aspx?type=B">�ڪ��ɥR����</a>
				</label>	-->	      
		  </div>		
		
		  <div class="Page">
        ��<asp:Label ID="PageNumberText" runat="server" Text="" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server" Text="" CssClass="Number" />���A
        �@<asp:Label ID="TotalRecordText" runat="server" Text="" CssClass="Number" />����ơA     
        <asp:HyperLink ID="PreviousLink" runat="server">
          <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg" Visible="false" AlternateText="�W�@��"></asp:image>        	            
          <asp:Label ID="PreviousText" runat="server" Visible="false" Text="Label">�W�@�� &nbsp;</asp:Label>            
        </asp:HyperLink>
        ��� <asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" /> �� &nbsp;
        <asp:HyperLink ID="NextLink" runat="server">
          <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg" Visible="false" AlternateText="�U�@��"></asp:image>                     	
          <asp:Label ID="NextText" runat="server" Visible="false" Text="Label">�U�@�� &nbsp;</asp:Label>
        </asp:HyperLink>�A�C�����                      
        <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
          <asp:ListItem Selected="True" Value="10">10</asp:ListItem>
          <asp:ListItem Value="20">20</asp:ListItem>
          <asp:ListItem Value="30">30</asp:ListItem>
          <asp:ListItem Value="50">50</asp:ListItem>
        </asp:DropDownList> ��
      </div>
      
      <!--�����C /�}�l -->
      <asp:Label ID="TableText" runat="server" Text="" />    
		
    	<!--�j�M�ڪ�����/�}�l -->
			<div class="mysearch">
			  <!--<label for="textfield">�j�M�ڤ��������D�G</label>
			  <asp:TextBox ID="KeywordText" runat="server" onfocus="this.value=''">�п�J����r</asp:TextBox>
        <asp:Button ID="SearchBtn" CssClass="btn2" runat="server" Text="�j�M" />
		    <asp:Button ID="AdSearchBtn" CssClass="btn2" runat="server" Text="�i���d��" Visible="false" />	-->	        		  
		  </div>	
	  </div>
	  </div>
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="�^�쭶���̤W��" /></a></div><!--20080829 chiung -->
  </td>	  

</asp:Content>


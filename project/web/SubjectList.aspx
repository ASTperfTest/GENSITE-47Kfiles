<%@ Page Language="VB" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="false"
    CodeFile="SubjectList.aspx.vb" Inherits="SubjectList" Title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <link rel="stylesheet" type="text/css" href="/css/jquery.autocomplete.css" />
    <script type="text/javascript" src="/js/jquery.autocomplete.js"></script>
    <script type="text/javascript" language="javascript">
        $(function() {
        $("#ctl00_ContentPlaceHolder1_t1").autocomplete('AutoComplete.aspx', { delay: 10 });
        });
	
    </script>
    
    <a title="�������e���" class="Accesskey" accesskey="C" href="#" xmlns:hyweb="urn:gip-hyweb-com">
        :::</a>
    <div class="path" xmlns:hyweb="urn:gip-hyweb-com">
        �ثe��m�G<a title="����" href="/mp.asp?mp=1">����</a>&gt;<a title="�D�D�]" href="#">�D�D�]</a></div>
    <h3>
        �D�D�]</h3>
    <div id="Magazine">
        <div class="Event">
            <div class="Page">
                ��<asp:label id="PageNumberText" runat="server" text="" cssclass="Number" />/
				<asp:label id="TotalPageText" runat="server" text="" cssclass="Number" />���A 
				
				�@<asp:label id="TotalRecordText" runat="server" text="" cssclass="Number" />����ơA
				
                <asp:hyperlink id="PreviousLink" runat="server">
        	        <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg" Visible="false" AlternateText="�W�@��"></asp:image>        	            
        	        <asp:Label ID="PreviousText" runat="server" Visible="false" Text="Label">�W�@�� &nbsp;</asp:Label>            
        	    </asp:hyperlink>
			
                ���
                <asp:dropdownlist id="PageNumberDDL" autopostback="true" runat="server" />
                ��</label>&nbsp;
				
                <asp:hyperlink id="NextLink" runat="server">
                    <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg" Visible="false" AlternateText="�U�@��"></asp:image>                     	
                    <asp:Label ID="NextText" runat="server" Visible="false" Text="Label">�U�@�� &nbsp;</asp:Label>
                </asp:hyperlink>
				
                �A �C�����
                <asp:dropdownlist id="PageSizeDDL" autopostback="true" runat="server">
                    <asp:ListItem Selected="True" Value="10">10</asp:ListItem>
                    <asp:ListItem Value="20">20</asp:ListItem>
                    <asp:ListItem Value="30">30</asp:ListItem>
                    <asp:ListItem Value="50">50</asp:ListItem>
                </asp:dropdownlist> ���A
				
				��
                <asp:DropDownList id="SortDDL" autopostback="true" runat="server" >
                    <asp:ListItem Selected="True" Value="0">�H��</asp:ListItem>
					<asp:ListItem Value="3">��s�ɶ�</asp:ListItem>
                    <asp:ListItem Value="1">�s������</asp:ListItem>
                    <asp:ListItem Value="2">���ƶq</asp:ListItem>
                </asp:DropDownList>
                �Ƨ�
            </div>
            <asp:label runat="server" id="TabText" text=""></asp:label>
            <asp:Label id="lblSearch" runat="server" text="�D�D�]�d�ߡG"></asp:Label>
            <asp:TextBox id="t1" runat="server" AutoCompleteType="Disabled"></asp:TextBox>
            <asp:Button id="btnSubmit" runat="server" text="�e�X" onclick="btnSubmit_Click" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <img src="/subject/xslGip/style1Q/images/search.gif" alt="search">
            <asp:HyperLink ID="HyperLink1" runat="server" Font-Size="Small" NavigateUrl="~/ArticleSearch.aspx">�峹�d��</asp:HyperLink>
            <asp:label id="TableText" runat="server" text=""></asp:label>
			
			<div class="Page">
                ��<asp:label id="PageNumberText_" runat="server" text="" cssclass="Number" />/
				<asp:label id="TotalPageText_" runat="server" text="" cssclass="Number" />���A 
				
				�@<asp:label id="TotalRecordText_" runat="server" text="" cssclass="Number" />����ơA
				
                <asp:hyperlink id="PreviousLink_" runat="server">
        	        <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg_" Visible="false" AlternateText="�W�@��"></asp:image>        	            
        	        <asp:Label ID="PreviousText_" runat="server" Visible="false" Text="Label">�W�@�� &nbsp;</asp:Label>            
        	    </asp:hyperlink>
			
                ���
                <asp:dropdownlist id="PageNumberDDL_" autopostback="true" runat="server" />
                ��</label>&nbsp;
				
                <asp:hyperlink id="NextLink_" runat="server">
                    <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg_" Visible="false" AlternateText="�U�@��"></asp:image>                     	
                    <asp:Label ID="NextText_" runat="server" Visible="false" Text="Label">�U�@�� &nbsp;</asp:Label>
                </asp:hyperlink>
				
                �A �C�����
                <asp:dropdownlist id="PageSizeDDL_" autopostback="true" runat="server">
                    <asp:ListItem Selected="True" Value="10">10</asp:ListItem>
                    <asp:ListItem Value="20">20</asp:ListItem>
                    <asp:ListItem Value="30">30</asp:ListItem>
                    <asp:ListItem Value="50">50</asp:ListItem>
                </asp:dropdownlist> ���C
				
            </div>
        </div>
    </div>
</asp:Content>



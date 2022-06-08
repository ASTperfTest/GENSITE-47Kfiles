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
    
    <a title="網頁內容資料" class="Accesskey" accesskey="C" href="#" xmlns:hyweb="urn:gip-hyweb-com">
        :::</a>
    <div class="path" xmlns:hyweb="urn:gip-hyweb-com">
        目前位置：<a title="首頁" href="/mp.asp?mp=1">首頁</a>&gt;<a title="主題館" href="#">主題館</a></div>
    <h3>
        主題館</h3>
    <div id="Magazine">
        <div class="Event">
            <div class="Page">
                第<asp:label id="PageNumberText" runat="server" text="" cssclass="Number" />/
				<asp:label id="TotalPageText" runat="server" text="" cssclass="Number" />頁， 
				
				共<asp:label id="TotalRecordText" runat="server" text="" cssclass="Number" />筆資料，
				
                <asp:hyperlink id="PreviousLink" runat="server">
        	        <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg" Visible="false" AlternateText="上一頁"></asp:image>        	            
        	        <asp:Label ID="PreviousText" runat="server" Visible="false" Text="Label">上一頁 &nbsp;</asp:Label>            
        	    </asp:hyperlink>
			
                到第
                <asp:dropdownlist id="PageNumberDDL" autopostback="true" runat="server" />
                頁</label>&nbsp;
				
                <asp:hyperlink id="NextLink" runat="server">
                    <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg" Visible="false" AlternateText="下一頁"></asp:image>                     	
                    <asp:Label ID="NextText" runat="server" Visible="false" Text="Label">下一頁 &nbsp;</asp:Label>
                </asp:hyperlink>
				
                ， 每頁顯示
                <asp:dropdownlist id="PageSizeDDL" autopostback="true" runat="server">
                    <asp:ListItem Selected="True" Value="10">10</asp:ListItem>
                    <asp:ListItem Value="20">20</asp:ListItem>
                    <asp:ListItem Value="30">30</asp:ListItem>
                    <asp:ListItem Value="50">50</asp:ListItem>
                </asp:dropdownlist> 筆，
				
				依
                <asp:DropDownList id="SortDDL" autopostback="true" runat="server" >
                    <asp:ListItem Selected="True" Value="0">隨機</asp:ListItem>
					<asp:ListItem Value="3">更新時間</asp:ListItem>
                    <asp:ListItem Value="1">瀏覽次數</asp:ListItem>
                    <asp:ListItem Value="2">文件數量</asp:ListItem>
                </asp:DropDownList>
                排序
            </div>
            <asp:label runat="server" id="TabText" text=""></asp:label>
            <asp:Label id="lblSearch" runat="server" text="主題館查詢："></asp:Label>
            <asp:TextBox id="t1" runat="server" AutoCompleteType="Disabled"></asp:TextBox>
            <asp:Button id="btnSubmit" runat="server" text="送出" onclick="btnSubmit_Click" />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <img src="/subject/xslGip/style1Q/images/search.gif" alt="search">
            <asp:HyperLink ID="HyperLink1" runat="server" Font-Size="Small" NavigateUrl="~/ArticleSearch.aspx">文章查詢</asp:HyperLink>
            <asp:label id="TableText" runat="server" text=""></asp:label>
			
			<div class="Page">
                第<asp:label id="PageNumberText_" runat="server" text="" cssclass="Number" />/
				<asp:label id="TotalPageText_" runat="server" text="" cssclass="Number" />頁， 
				
				共<asp:label id="TotalRecordText_" runat="server" text="" cssclass="Number" />筆資料，
				
                <asp:hyperlink id="PreviousLink_" runat="server">
        	        <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg_" Visible="false" AlternateText="上一頁"></asp:image>        	            
        	        <asp:Label ID="PreviousText_" runat="server" Visible="false" Text="Label">上一頁 &nbsp;</asp:Label>            
        	    </asp:hyperlink>
			
                到第
                <asp:dropdownlist id="PageNumberDDL_" autopostback="true" runat="server" />
                頁</label>&nbsp;
				
                <asp:hyperlink id="NextLink_" runat="server">
                    <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg_" Visible="false" AlternateText="下一頁"></asp:image>                     	
                    <asp:Label ID="NextText_" runat="server" Visible="false" Text="Label">下一頁 &nbsp;</asp:Label>
                </asp:hyperlink>
				
                ， 每頁顯示
                <asp:dropdownlist id="PageSizeDDL_" autopostback="true" runat="server">
                    <asp:ListItem Selected="True" Value="10">10</asp:ListItem>
                    <asp:ListItem Value="20">20</asp:ListItem>
                    <asp:ListItem Value="30">30</asp:ListItem>
                    <asp:ListItem Value="50">50</asp:ListItem>
                </asp:dropdownlist> 筆。
				
            </div>
        </div>
    </div>
</asp:Content>



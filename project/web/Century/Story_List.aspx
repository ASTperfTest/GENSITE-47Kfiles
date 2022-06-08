<%@ Page Title="" Language="C#" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true"
    CodeFile="Story_List.aspx.cs" Inherits="Century_Story_List" %>

<%@ Register Src="TabText.ascx" TagName="TabMenu" TagPrefix="uc1" %>
<%@ Register Src="path.ascx" TagName="path" TagPrefix="uc2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
<style type="text/css">
    #List img{
	    float:left;
	    margin: 0 8px 4px 0;
	    width:160px;
	    height:auto;
	    border: 1px solid #fff;
    }
    #List a:hover img
    {
	    border: 1px solid #888;
    }   
    #List a
    {
        color: #0b5891;
        text-decoration: underline;
        padding-top: 0px;
        padding-right: 0px;
        padding-left: 0px;
        padding-bottom: 5px;
        font-size:1.1em;
        margin: 0px 1em 5px 0px;
    }
    #List a:hover
    {
        color: #008a06;
	    text-decoration: none;
    }
    
    #xBody
    {
        margin: 0.5em 0px 0px;
        display: block;
    }
</style>
    <!--中間內容區-->
    <!--  目前位置  -->
    <uc2:path runat="server" ID="path1" />
    <%--<div class="path" style="float: left; padding-top: 10px">
        目前位置：
    </div>
    <div style="float: left; width: 90%">
        <ul id="path_menu">
            <li><a href="../mp.asp?mp=1">首頁</a></li>
            <li style='top: 10px;'>></li>
            <li><a href="Events_List.aspx">農業百年發展史</a></li>
        </ul>
    </div>--%>
    <!--標籤區-->
    <uc1:TabMenu runat="server" ID="TabMenu1" />
    <!--文章列表-->
    <div id="Event">
        <!-- 分頁列表  Start -->
        <div class="Page">
            第<asp:Label ID="PageNumberText" runat="server" Text="" CssClass="Number" />/<asp:Label
                ID="TotalPageText" runat="server" Text="" CssClass="Number" />頁， 共<asp:Label ID="TotalRecordText"
                    runat="server" Text="" CssClass="Number" />筆資料，
            <asp:LinkButton ID="PreviousLink" runat="server" OnClick="PreviousLink_Click" Visible="false">
                <asp:Image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg"
                    Visible="false" AlternateText="上一頁"></asp:Image>
                <asp:Label ID="PreviousText" runat="server">上一頁</asp:Label>
            </asp:LinkButton>
            &nbsp; 到第
            <asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" />
            頁 &nbsp;
            <asp:LinkButton ID="NextLink" runat="server" OnClick="NextLink_Click" Visible="false">
                <asp:Image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg"
                    Visible="false" AlternateText="下一頁"></asp:Image>
                <asp:Label ID="NextText" runat="server">下一頁</asp:Label>
            </asp:LinkButton>
            &nbsp;，每頁顯示
            <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
                <asp:ListItem Selected="True" Value="10">10</asp:ListItem>
                <asp:ListItem Value="20">20</asp:ListItem>
                <asp:ListItem Value="30">30</asp:ListItem>
                <asp:ListItem Value="50">50</asp:ListItem>
            </asp:DropDownList>
            筆
        </div>
        <!-- 分頁列表   End  -->
        <!-- 文章列表  Start -->
        <div id="List" class="lp">
            <asp:Repeater runat="server" ID="rptList" OnItemDataBound="rptList_ItemDataBound">
                <HeaderTemplate>
                    <table>
                </HeaderTemplate>
                <ItemTemplate>
                    <tr>
                        <td>
                            <asp:Label runat="server" ID="labICUItem" Visible="false" Text='<%# Eval("iCUItem") %>'></asp:Label>
                            <a title="" href="#"><asp:Image runat="server" ID="imgArticle" ImageUrl='<%# Eval("xImgFile") %>' /></a>
                        </td>
                        <td style="vertical-align:top;line-height:150%;padding-top:15px;">
                            <asp:HyperLink runat="server" ID="HyperLink1" Text='<%# Eval("sTitle")%>'></asp:HyperLink>
                            <asp:Label runat="server" ID="labDate" Text='<%# Eval("dEditDate")%>'></asp:Label>
                            <span id="xBody"><asp:Label runat="server" ID="labContent" Text='<%# Eval("xBody")%>'></asp:Label></span>
                        </td>
                    </tr>
                </ItemTemplate>
                <FooterTemplate>
                    </table>
                </FooterTemplate>
            </asp:Repeater>
        </div>
        <!-- 文章列表   End  -->
    </div>
</asp:Content>

<%@ Page Title="" Language="C#" MasterPageFile="~/default.master" AutoEventWireup="true"
    CodeFile="categorylist.aspx.cs" Inherits="Category_categorylist" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <td class="kleft">
        <span class="kleftStyle">
            <asp:TreeView ID="CatTreeView" runat="server" NodeIndent="5" ImageSet="BulletedList4"
                ShowExpandCollapse="False">
                <ParentNodeStyle Font-Bold="False" />
                <SelectedNodeStyle Font-Underline="True" ForeColor="#5555DD" HorizontalPadding="5"
                    VerticalPadding="2" />
                <NodeStyle Font-Names="Verdana" Font-Size="8pt" ForeColor="Black" HorizontalPadding="5"
                    NodeSpacing="1" VerticalPadding="2" />
                <DataBindings>
                    <asp:TreeNodeBinding DataMember="node" TextField="name" NavigateUrlField="url" />
                </DataBindings>
            </asp:TreeView>
            <asp:XmlDataSource ID="CatTreeViewDS" EnableCaching="false" runat="server" />
        </span>
    </td>
    <td class="center">
        <a title="網頁內容資料" class="Accesskey" accesskey="C" href="#" xmlns:hyweb="urn:gip-hyweb-com">
            :::</a>
        <asp:Label runat="server" ID="TabText" Text=""></asp:Label>
        <div class="path" xmlns:hyweb="urn:gip-hyweb-com">
            目前位置：<a title="首頁" href="/mp.asp?mp=1">首頁</a>&gt;
            <asp:Label runat="server" ID="NavUrlText" Text=""></asp:Label>
        </div>
        <h3>
            <asp:Label runat="server" ID="NavTitleText" Text=""></asp:Label>
        </h3>
        <div id="Magazine">
            <div class="Event">
                <asp:Label ID="NodeText" runat="server" Text="" />
                <div class="Page" id="pageDIV" runat="server">
                    第<asp:Label ID="PageNumberText" runat="server" Text="" CssClass="Number" />/<asp:Label
                        ID="TotalPageText" runat="server" Text="" CssClass="Number" />頁， 共<asp:Label ID="TotalRecordText"
                            runat="server" Text="" CssClass="Number" />筆資料，
                    <asp:HyperLink ID="PreviousLink" runat="server">
                        <asp:Image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg"
                            Visible="false" AlternateText="上一頁"></asp:Image>
                        <asp:Label ID="PreviousText" runat="server" Visible="false" Text="Label">上一頁 &nbsp;</asp:Label>
                    </asp:HyperLink>
                    到第
                    <asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" />
                    頁</label>&nbsp;
                    <asp:HyperLink ID="NextLink" runat="server">
                        <asp:Image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg"
                            Visible="false" AlternateText="下一頁"></asp:Image>
                        <asp:Label ID="NextText" runat="server" Visible="false" Text="Label">下一頁 &nbsp;</asp:Label>
                    </asp:HyperLink>
                    ，每頁顯示
                    <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
                        <asp:ListItem Selected="True" Value="10">10</asp:ListItem>
                        <asp:ListItem Value="20">20</asp:ListItem>
                        <asp:ListItem Value="30">30</asp:ListItem>
                        <asp:ListItem Value="50">50</asp:ListItem>
                    </asp:DropDownList>
                    筆
                </div>
                <asp:Label ID="TableText" runat="server" Text="" />
            </div>
        </div>
    </td>
</asp:Content>

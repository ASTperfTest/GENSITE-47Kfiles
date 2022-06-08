<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DsdASPXList.aspx.cs" Inherits="DsdASPXList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>資料上稿-入口網</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <div id="FuncName">
            <h1>
                資料管理／資料上稿-入口網<font size="2">【目錄樹節點: <%=Session["catName"]%>】</font>
            </h1>
            <div id="Nav">
                <a href="DsdASPXAdd.aspx" title="新增">新增</a>
                <a href="DsdASPXQuery.aspx" title="查詢">查詢</a>
            </div>
            <div id="ClearFloat">
            </div>
        </div>
        <div id="FormName">
            單元資料維護【主題單元:<%=Session["CtUnitName"]%> / 單元資料:純網頁】 
        </div>
        　<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
            <tr>
                <td width="95%" colspan="2" height="230" valign="top">
                    <div align="center">
                        共
                        <asp:Label ID="TotalRecordText" runat="server" Text="" CssClass="Number" />
                        筆資料，&nbsp;每頁顯示:&nbsp;
                        <asp:DropDownList ID="PageSizeDDL" Style="color: #FF0000" AutoPostBack="true" runat="server"
                            OnSelectedIndexChanged="PageSizeDDL_SelectedIndexChanged">
                            <asp:ListItem Value="15">15</asp:ListItem>
                            <asp:ListItem Value="30">30</asp:ListItem>
                            <asp:ListItem Value="50">50</asp:ListItem>
                            <asp:ListItem Value="300">300</asp:ListItem>
                        </asp:DropDownList>筆，
                        <asp:LinkButton ID="PreviousLink" runat="server" Visible="false" OnClick="PreviousLink_Click">
                            <asp:Image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg"
                                Visible="false" AlternateText="上一頁"></asp:Image>
                            <asp:Label ID="PreviousText" runat="server">上一頁</asp:Label>
                        </asp:LinkButton>&nbsp;目前在第
                        <asp:DropDownList ID="PageNumberDDL" Style="color: #FF0000;" AutoPostBack="true"
                            runat="server" OnSelectedIndexChanged="PageNumberDDL_SelectedIndexChanged" />
                        &nbsp;頁&nbsp;
                        <asp:LinkButton ID="NextLink" runat="server" Visible="false" OnClick="NextLink_Click">
                            <asp:Image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg"
                                Visible="false" AlternateText="下一頁"></asp:Image>
                            <asp:Label ID="NextText" runat="server">下一頁</asp:Label>
                        </asp:LinkButton>
                        <asp:Repeater ID="rptList" runat="server" OnItemDataBound="rptList_ItemDataBound">
                            <HeaderTemplate>
                                <table id="ListTable">
                                    <tr>
                                        <th>
                                            預覽
                                        </th>
                                        <th>
                                            標題
                                        </th>
                                        <th>
                                            檔案下載
                                        </th>
                                        <th>
                                            是否公開
                                        </th>
                                        <th>
                                            圖檔
                                        </th>
                                    </tr>
                            </HeaderTemplate>
                            <ItemTemplate>
                                <tr>
                                    <td class="eTableContent" style="text-align: center; width: 15%;">
                                       <asp:Literal runat="server" ID="litiCUItem" Text='<%# Eval("iCUItem") %>' Visible="false"></asp:Literal>
                                       <%--<a href="#" target="_blank">View</a>--%>
                                       <asp:HyperLink runat="server" ID="linkView" Text='View' Target="_blank"></asp:HyperLink>
                                    </td>
                                    <td class="eTableContent" style="text-align: center; width: 15%;">
                                       <%--<a href="#"><asp:Literal runat="server" ID="litTitle" Text='<%# Eval("sTitle") %>'></asp:Literal></a>--%>
                                       <asp:HyperLink runat="server" ID="linkEdit" Text='<%# Eval("sTitle") %>'></asp:HyperLink>
                                    </td>
                                    <td class="eTableContent" style="text-align: center; width: 20%;">
                                       <asp:Literal runat="server" ID="litFileName" Text='<%# Eval("fileDownLoad")%>'></asp:Literal>
                                    </td>
                                    <td class="eTableContent" style="width: 20%;">
                                       <asp:Literal runat="server" ID="litPublic" Text='<%# Eval("fCTUPublic")%>'></asp:Literal>
                                    </td>
                                    <td class="eTableContent" style="width: 20%;">
                                       <asp:Literal runat="server" ID="litPicture" Text='<%# Eval("xImgfile") %>'></asp:Literal>
                                    </td>
                                </tr>
                            </ItemTemplate>
                            <FooterTemplate>
                                </table>
                            </FooterTemplate>
                        </asp:Repeater>
                    </div>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>

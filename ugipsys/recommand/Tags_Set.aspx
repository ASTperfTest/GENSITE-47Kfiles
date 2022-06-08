<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Tags_Set.aspx.cs" Inherits="Tags_Set" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>功能管理 / 好文推薦 / 設定分類</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <script language="javascript" type="text/javascript">
        // 實作全選
        function Select_All(spanChk) {
            elm = document.forms[0];
            for (i = 0; i <= elm.length - 1; i++) {
                if (elm[i].type == "checkbox" && elm[i].id != spanChk.id) {
                    if (!elm.elements[i].checked)
                        elm.elements[i].click();
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <div id="FuncName">
            <h1>好文推薦-分類管理</h1>
            <div id="Nav">
            </div>
            <div id="ClearFloat">
            </div>
            <div id="Event">
                <br />
                分類：<asp:TextBox runat="server" ID="txtSearch"></asp:TextBox>
                <asp:Button runat="server" ID="btnAdd" CssClass="cbutton" Text="新增" OnClick="btnAdd_Click"
                    OnClientClick="return confirm('確認新增?');" />
                <asp:Button runat="server" ID="btnSearch" CssClass="cbutton" Text="搜尋" OnClick="btnSearch_Click" />
                <br />
                <br />
            </div>
        </div>
        <div id="FormName">
        </div>
        <table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
            <tr>
                <td width="95%" colspan="2" height="230" valign="top">
                    <div>
                        共&nbsp;<asp:Label ID="TotalRecordText" runat="server" Text="" CssClass="Number" style="color:#FF0000" />&nbsp;筆資料，
                        &nbsp;每頁顯示&nbsp;<font color="#FF0000">10</font>&nbsp;筆，目前在第
                        <asp:DropDownList ID="PageNumberDDL" Style="color: #FF0000" AutoPostBack="true" 
                            runat="server" onselectedindexchanged="PageNumberDDL_SelectedIndexChanged" />
                        &nbsp;頁&nbsp;
                        <asp:LinkButton ID="PreviousLink" runat="server" Visible="false" onclick="PreviousLink_Click" >
                            <asp:Image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg"
                                Visible="false" AlternateText="上一頁"></asp:Image>
                            <asp:Label ID="PreviousText" runat="server">上一頁</asp:Label>
                        </asp:LinkButton>&nbsp;
                        <asp:LinkButton ID="NextLink" runat="server" Visible="false" onclick="NextLink_Click" >
                            <asp:Image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg"
                                Visible="false" AlternateText="下一頁"></asp:Image>
                            <asp:Label ID="NextText" runat="server">下一頁</asp:Label>
                        </asp:LinkButton>
                        <asp:Repeater ID="rptList" runat="server" 
                            onitemdatabound="rptList_ItemDataBound" >
                            <HeaderTemplate>
                                <table cellspacing="0" id="ListTable">
                                    <tr>
                                        <th style="width: 100px">
                                            <input type="button" name="btnSelect" class="cbutton" value="全選" onclick="Select_All(this);" />
                                        </th>
                                        <th>
                                            分類名稱
                                        </th>
                                        <th style="width:100px">
                                            分類文章數
                                        </th>
                                    </tr>
                            </HeaderTemplate>
                            <ItemTemplate>
                                <tr>
                                    <td class="eTableContent" style="text-align: center;">
                                        <asp:CheckBox runat="server" ID="checkbox1" />
                                    </td>
                                    <td align="center" class="eTableContent">
                                        <asp:Label runat="server" ID="labtagID" Text='<%# Eval("tagID") %>' Visible="false"></asp:Label>
                                        <asp:Label runat="server" ID="labTAGs" Text='<%# Eval("tagName") %>'></asp:Label>
                                    </td>
                                    <td style="text-align: center;">
                                        <asp:Label runat="server" ID="labCount" Text='<%# Eval("intCount")%>'></asp:Label>
                                    </td>
                                </tr>
                            </ItemTemplate>
                            <FooterTemplate>
                                </table></FooterTemplate>
                        </asp:Repeater>
                    </div>
                </td>
            </tr>
            <tr>
                <td style="text-align: center;">
                    <asp:Button runat="server" ID="btnDelete" CssClass="cbutton" Text="刪除" OnClick="btnDelete_Click" OnClientClick="return confirm('確定刪除?');" />
                    <asp:Button runat="server" ID="btnConfirm" CssClass="cbutton" Text="關閉視窗" OnClientClick="javascript:void(window.close());" />
            </tr>
        </table>
    </div>
    </form>
</body>
<%--<script type="text/javascript">
    function page(p) {
        window.location.href = 'Index.aspx?sTitle=<%=Server.UrlEncode(titles)%>&Status=<%=status%>&Types=<%=types%>&pagesize=<%=pl.PageSize%>&page=' + (p - 1);
    }
    function pagesize(ps) {
        window.location.href = 'Index.aspx?sTitle=<%=Server.UrlEncode(titles)%>&Status=<%=status%>&Types=<%=types%>&pagesize=' + ps;
    }
</script>--%>
</html>

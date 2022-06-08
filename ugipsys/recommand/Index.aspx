<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Index.aspx.cs" Inherits="Index" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>報馬仔文章管理</title>
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
            <h1>
                資料管理／報馬仔文章管理</h1>
            <div id="Nav">
                <a href="#" title="分類管理" onclick="javascript:void(window.open('Tags_Set.aspx','','width=500px,height=600px,toolbar=no,location=no'));">分類管理</a>
            </div>
            <div id="ClearFloat">
            </div>
        </div>
        <div id="FormName">
            <table>
                <tr>
                    <td>
                        條件篩選依：審核狀態
                    </td>
                    <td>
                        <asp:DropDownList runat="server" ID="ddlStatus" AutoPostBack="true" 
                            onselectedindexchanged="ddlStatus_SelectedIndexChanged">
                            <asp:ListItem Text="全部" Value="0"></asp:ListItem>
                            <asp:ListItem Text="尚未審核" Value="1" Selected="True"></asp:ListItem>
                            <asp:ListItem Text="已通過" Value="2"></asp:ListItem>
                            <asp:ListItem Text="未通過" Value="3"></asp:ListItem>
                        </asp:DropDownList>
                    </td>
                    <td>&nbsp;&nbsp;&nbsp;</td>
                    <td>關鍵字：</td>
                    <td><asp:TextBox runat="server" ID="txtKeyword" Text=""></asp:TextBox></td>
                    <td>&nbsp;&nbsp;&nbsp;</td>
                    <td><asp:Button runat="server" ID="btnSearch" Text="搜尋" onclick="btnSearch_Click" /></td>
                </tr>
            </table>
        </div>
        <table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
            <tr>
                <td width="95%" colspan="2" height="230" valign="top">
                    <div align="center">
                        第
                        <asp:Label ID="PageNumberText" runat="server" Text="" CssClass="Number" />
                        /
                        <asp:Label ID="TotalPageText" runat="server" Text="" CssClass="Number" />
                        頁 ， 共
                        <asp:Label ID="TotalRecordText" runat="server" Text="" CssClass="Number" />
                        筆資料，
                        <asp:LinkButton ID="PreviousLink" runat="server" Visible="false" OnClick="PreviousLink_Click">
                            <asp:Image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg"
                                Visible="false" AlternateText="上一頁"></asp:Image>
                            <asp:Label ID="PreviousText" runat="server">上一頁</asp:Label>
                        </asp:LinkButton>&nbsp;跳至第
                        <asp:DropDownList ID="PageNumberDDL" Style="color: #FF0000;" AutoPostBack="true"
                            runat="server" OnSelectedIndexChanged="PageNumberDDL_SelectedIndexChanged" />
                        &nbsp;頁&nbsp;
                        <asp:LinkButton ID="NextLink" runat="server" Visible="false" OnClick="NextLink_Click">
                            <asp:Image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg"
                                Visible="false" AlternateText="下一頁"></asp:Image>
                            <asp:Label ID="NextText" runat="server">下一頁</asp:Label>
                        </asp:LinkButton>，&nbsp;每頁筆數:&nbsp;
                        <asp:DropDownList ID="PageSizeDDL" Style="color: #FF0000" AutoPostBack="true" runat="server"
                            OnSelectedIndexChanged="PageSizeDDL_SelectedIndexChanged">
                            <asp:ListItem Value="10">10</asp:ListItem>
                            <asp:ListItem Value="20">20</asp:ListItem>
                            <asp:ListItem Value="30">30</asp:ListItem>
                            <asp:ListItem Value="50">50</asp:ListItem>
                        </asp:DropDownList>
                        <asp:Repeater ID="rptList" runat="server" OnItemDataBound="rptList_ItemDataBound">
                            <HeaderTemplate>
                                <table id="ListTable">
                                    <tr>
                                        <th>
                                            <input type="button" name="btnSelectAll" value="全選" onclick="Select_All(this)" />
                                        </th>

                                        <th>
                                            編修
                                        </th>                                        
                                        <th>
                                            審核
                                        </th>
                                        <th>
                                            推薦日期
                                        </th>
                                        <th>
                                            推薦文章標題
                                        </th>
                                        <th>
                                            推薦原因
                                        </th>
                                        <th>
                                            分類
                                        </th>
                                    </tr>
                            </HeaderTemplate>
                            <ItemTemplate>
                                <tr>
                                    <td class="eTableContent" style="text-align: center; width: 5%;">
                                        <asp:CheckBox runat="server" ID="checkbox1" />
                                        <%--序號--%>
                                        <%--<%# ((Convert.ToInt32(PageNumberDDL.SelectedValue) - 1) * Convert.ToInt32(PageSizeDDL.SelectedValue)) + (Container.ItemIndex +1) %>--%>
                                        <asp:Label runat="server" ID="labcID" Text='<%# Eval("cID") %>' Visible="false"></asp:Label>
                                    </td>

                                    <td class="eTableContent" style="text-align: center; width: 5%;">
                                        <asp:HyperLink runat="server" ID="HyperLink1" Text="編修"></asp:HyperLink>
                                    </td>                                    
                                    <td class="eTableContent" style="text-align: center; width:60px;">
                                        <asp:Label runat="server" ID="labStatus" Text='<%# Eval("Status") %>'></asp:Label>
                                    </td>
                                    <td class="eTableContent" style="text-align: center; width:80px;">
                                        <asp:Label runat="server" ID="labDate" Text='<%# Eval("Created_Date") %>'></asp:Label>
                                    </td>
                                    <td class="eTableContent" style="width: 10%;">
										<asp:HyperLink NavigateUrl='<%# Eval("url") %>'
										Text='<%# Eval("Title") %>'
										Target="_new"
										runat="server" />
                                    </td>
                                    <td class="eTableContent">
                                        <asp:Label runat="server" ID="labContent" Text='<%# Eval("aContent") %>'></asp:Label>
                                        <p>(推薦會員：<%# Eval("account") %>/<%# Eval("nickname") %>/<%# Eval("realname") %>)</p>
                                    </td>
                                    <td class="eTableContent" style="text-align: center; width: 5%;">
                                        <asp:Label runat="server" ID="labTAGs"></asp:Label>
                                    </td>
                                </tr>
                            </ItemTemplate>
                            <FooterTemplate>
                                <tr>
                                    <td style="text-align: center;" colspan="7">
                                        <asp:Button runat="server" ID="btnYes" CssClass="cbutton" Text="通過" OnClick="btnYes_Click" />&nbsp;&nbsp;<asp:Button
                                            runat="server" ID="btnNo" CssClass="cbutton" Text="未通過" OnClick="btnNo_Click" />
                                    </td>
                                </tr>
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

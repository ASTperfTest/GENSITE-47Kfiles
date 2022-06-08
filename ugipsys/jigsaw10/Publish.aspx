<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Publish.aspx.cs" Inherits="Publish" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>農漁生產作物管理</title>
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <link type="text/css" rel="stylesheet" href="../css/setstyle.css" />
</head>
<body>
    <form id="form1" runat="server" enableviewstate="False">
    <div>
        <table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
            <tr>
                <td width="50%" align="left" nowrap="nowrap" class="FormName">
                    農漁生產作物管理&nbsp; <font size="2">
                        <asp:Literal ID="titles" runat="server"></asp:Literal></font>
                </td>
                <td width="50%" class="FormLink" align="right">
                    <a href="<%=returnUrl%>" title="回前頁">回前頁</a>
                </td>
            </tr>
            <tr>
                <td width="100%" colspan="2">
                    <hr />
                    <br />
                </td>
            </tr>
            <tr>
                <td width="95%" colspan="2" height="230" valign="top">
                    <div align="center">
                        <asp:Repeater ID="rptList" runat="server">
                            <HeaderTemplate>
                                <table id="ListTable" width="100%" cellpadding="0" cellspacing="1">
                                    <tr align="left">
                                        <th>
                                            專區區塊條列標題
                                        </th>
                                        <th width="15%">
                                            是否公開
                                        </th>
                                        <th width="25%">
                                            內容文章設定
                                        </th>
                                    </tr>
                            </HeaderTemplate>
                            <ItemTemplate>
                                <tr>
                                    <%# Eval("contextTitle")%>
                                    <%# Eval("contextOrder")%>
                                    <%# Eval("contextPublic")%>
                                    <%# Eval("contextAction")%>
                                </tr>
                            </ItemTemplate>
                            <FooterTemplate>
                                </table>
                                <div align="center">
                                    <input type="submit" value="編修存檔" class="cbutton" />
                                    <input type="reset" value="重　填" class="cbutton" />
                                    <input type="button" value="回前頁" class="cbutton" onclick="location.href='<%=returnUrl%>'" />
                                </div>
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

<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Publish.aspx.cs" Inherits="Publish" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>�A���Ͳ��@���޲z</title>
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
                    �A���Ͳ��@���޲z&nbsp; <font size="2">
                        <asp:Literal ID="titles" runat="server"></asp:Literal></font>
                </td>
                <td width="50%" class="FormLink" align="right">
                    <a href="<%=returnUrl%>" title="�^�e��">�^�e��</a>
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
                                            �M�ϰ϶����C���D
                                        </th>
                                        <th width="15%">
                                            �O�_���}
                                        </th>
                                        <th width="25%">
                                            ���e�峹�]�w
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
                                    <input type="submit" value="�s�צs��" class="cbutton" />
                                    <input type="reset" value="���@��" class="cbutton" />
                                    <input type="button" value="�^�e��" class="cbutton" onclick="location.href='<%=returnUrl%>'" />
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

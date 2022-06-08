<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DsdASPXEdit.aspx.cs" Inherits="DsdASPXEdit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
        <title>資料上稿-入口網</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <link type="text/css" rel="stylesheet" href="../css/form.css" />
</head>
<body>
    <form id="Form1" runat="server">
    <div>
        <div id="FuncName">
            <h1>
                資料管理／資料上稿-入口網<font size="2">【目錄樹節點: <%=Session["catName"]%>】</font>
            </h1>
            <div id="Nav">
                <a href="DsdASPXAdd.aspx?ItemID=<%=Session["itemID"] %>&CtNodeID=<%=Session["ctNodeId"] %>"
                    title="新增">新增</a> <a href="DsdASPXList.aspx?ItemID=<%=Session["itemID"] %>&CtNodeID=<%=Session["ctNodeId"] %>"
                        title="回條列">回條列</a>
            </div>
            <div id="ClearFloat">
            </div>
        </div>
        <div id="FormName">
            單元資料維護【主題單元:<%=Session["CtUnitName"]%>
            / 單元資料:純網頁】
        </div>
        <div>
            <table style="margin-top: 10px">
                <%--<tr>
                <td align="right" class="Label">檔案下載</td>
                <td><asp:FileUpload runat="server" ID="file1" /></td>
                </tr>--%>
                <tr>
                    <td align="right" class="Label">
                        <font color="red">*</font>標題
                    </td>
                    <td>
                        <asp:TextBox runat="server" ID="txtTitle" Width="250px"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="right" class="Label">
                        張貼日
                    </td>
                    <td>
                        <asp:TextBox runat="server" ID="txtDate" Width="75px"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="right" class="Label">
                        <font color="red">*</font>是否公開
                    </td>
                    <td>
                        <asp:DropDownList runat="server" ID="ddlPublic">
                            <asp:ListItem>請選擇</asp:ListItem>
                            <asp:ListItem Value="Y" Selected="True">公開</asp:ListItem>
                            <asp:ListItem Value="N">不公開</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td align="right" class="Label">
                        重要性
                    </td>
                    <td>
                        <asp:TextBox runat="server" ID="txtImport" Width="50px"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="right" class="Label">
                        <font color="red">*</font>單位
                    </td>
                    <td>
                        <asp:DropDownList runat="server" ID="ddlUnit">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td align="right" class="Label">
                        圖檔
                    </td>
                    <td>
                        <asp:FileUpload runat="server" ID="fileuploadImg" /><br />
                        <asp:Image runat="server" ID="imageNow" />
                    </td>
                </tr>
            </table>
        </div>
        <table>
            <tr>
                <td width="100%">
                    <asp:Button runat="server" ID="btnConfirm" CssClass="cbutton" Text="編修存檔" OnClick="btnConfirm_Click" />
                    <asp:Button runat="server" ID="btnDelete" CssClass="cbutton" Text="刪　　除" OnClick="btnDelete_Click" 
                                OnClientClick="return confirm('確定刪除?');" />
                    <input type="reset" id="btnRest" class="cbutton" value="重填" />
                    <input type="button" id="btnBack" class="cbutton" value="回前頁" onclick="javascript:void(history.back());" />
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>

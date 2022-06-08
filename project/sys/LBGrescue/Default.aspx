<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="rescure._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
    <style type="text/css">
        .style1
        {
            width: 137px;
        }
        .style2
        {
            width: 408px;
        }
        .style4
        {
            width: 594px;
        }
        .style7
        {
            width: 495px;
        }
        .style8
        {
            width: 227px;
        }
        .style9
        {
            width: 220px;
        }
        .style10
        {
            width: 236px;
        }
        .style11
        {
            width: 132px;
        }
        .style13
        {
            width: 234px;
        }
        .style14
        {
            width: 126px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:TextBox ID="txtUserID" runat="server"></asp:TextBox>
        <asp:Button ID="btn_Search" runat="server" onclick="btn_Search_Click" 
            Text="搜尋" />
        <br />
        <asp:Panel ID="Panel2" runat="server" GroupingText="菜葉甘藷" Width="926px">
            <table border="1" cellpadding="0" cellspacing="0" style="width: 96%;">
                <tr>
                    <td class="style10" rowspan="3">
                        <asp:Button ID="btn_Fix1" runat="server" onclick="btn_Fix1_Click" Text="執行修正" />
                    </td>
                    <td class="style2">
                        目前狀態</td>
                    <td class="style4">
                        <asp:Label ID="lbl_STATUS1" runat="server" Text="Label"></asp:Label>
                    </td>
                    <td class="style4">
                        &nbsp;</td>
                </tr>
                <tr>
                    <td class="style2">
                        成長天數(DAYS)</td>
                    <td class="style4">
                        <asp:TextBox ID="txtDAY_1" runat="server"></asp:TextBox>
                    </td>
                    <td class="style4">
                        &nbsp;</td>
                </tr>
                <tr>
                    <td class="style2">
                        實際天(GROWUPDAYS)</td>
                    <td class="style4">
                        <asp:TextBox ID="txtGDAY_1" runat="server"></asp:TextBox>
                    </td>
                    <td class="style4">
                        &nbsp;</td>
                </tr>
            </table>
            <asp:HiddenField ID="hf_UID1" runat="server" />
            <br />
        </asp:Panel>
        <asp:Panel ID="Panel3" runat="server" GroupingText="孤挺花" Width="923px" 
            Height="101px">
            <table border="1" cellpadding="0" cellspacing="0" 
                style="width: 96%; height: 52px;">
                <tr>
                    <td class="style1" rowspan="3">
                        <asp:Button ID="btn_Fix2" runat="server" onclick="btn_Fix2_Click" Text="執行修正" />
                        <asp:HiddenField ID="hf_UID2" runat="server" />
                    </td>
                    <td class="style9">
                        目前狀態</td>
                    <td>
                        <asp:Label ID="lbl_STATUS2" runat="server" Text="Label"></asp:Label>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td class="style9">
                        成長天數(DAYS)</td>
                    <td>
                        <asp:TextBox ID="txtDAY_2" runat="server"></asp:TextBox>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td class="style9">
                        實際天(GROWUPDAYS)</td>
                    <td>
                        <asp:TextBox ID="txtGDAY_2" runat="server"></asp:TextBox>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="Panel4" runat="server" GroupingText="毛豆" Width="926px">
            <table border="1" cellpadding="0" cellspacing="0" style="width: 96%;">
                <tr>
                    <td class="style11" rowspan="3">
                        <asp:Button ID="btn_Fix3" runat="server" onclick="btn_Fix3_Click" Text="執行修正" />
                    </td>
                    <td class="style8">
                        目前狀態</td>
                    <td>
                        <asp:Label ID="lbl_STATUS3" runat="server"></asp:Label>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td class="style8">
                        成長天數(DAYS)</td>
                    <td>
                        <asp:TextBox ID="txtDAY_3" runat="server" style="margin-left: 0px"></asp:TextBox>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
                <tr>
                    <td class="style8">
                        實際天(GROWUPDAYS)</td>
                    <td>
                        <asp:TextBox ID="txtGDAY_3" runat="server"></asp:TextBox>
                    </td>
                    <td>
                        &nbsp;</td>
                </tr>
            </table>
            <asp:HiddenField ID="hf_UID3" runat="server" />
            <br />
        </asp:Panel>
        <asp:Panel ID="Panel1" runat="server" GroupingText="彩色海芋" Width="923px">
            <table border="1" cellpadding="0" cellspacing="0" style="width: 96%;">
                <tr>
                    <td class="style14" rowspan="3">
                        <asp:Button ID="btn_Fix4" runat="server" onclick="btn_Fix4_Click" Text="執行修正" />
                    </td>
                    <td class="style13">
                        目前狀態</td>
                    <td class="style7">
                        <asp:Label ID="lbl_STATUS4" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="style13">
                        成長天數(DAYS)</td>
                    <td class="style7">
                        <asp:TextBox ID="txtDAY_4" runat="server"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="style13">
                        實際天(GROWUPDAYS)</td>
                    <td class="style7">
                        <asp:TextBox ID="txtGDAY_4" runat="server"></asp:TextBox>
                    </td>
                </tr>
            </table>
            <asp:HiddenField ID="hf_UID4" runat="server" />
            <br />
        </asp:Panel>
    
    </div>
    </form>
</body>
</html>

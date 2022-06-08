<%@ Control Language="C#" AutoEventWireup="true" CodeFile="EntryControl.ascx.cs" Inherits="UserControls_EntryControl" %>
<div id="divHide" runat="server" class="menu">
    <table>
        <tr>
            <td>
                <a id="Hide" runat="server" href="">隱藏</a> </td>
        </tr>
    </table>
</div>
<div id="divApprove" runat="server" class="menu">
    <table>
        <tr>
            <td>
                <a id="Approve" runat="server" href="">開啟</a> </td>
        </tr>
    </table>
</div>
<div id="divE" runat="server" class="menu">
    <table>
        <tr>
            <td>
                <a id="Edit" runat="server" href="">編輯</a> </td>
            <td>
                <a id="Delete" runat="server" href="">刪除</a> </td>
        </tr>
    </table>
</div>
<div id="data" runat="server">
    <table class="grid">
        <tr>
            <td align="left">
                <asp:Label ID="LabelDate" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <td align="left">
                <asp:Label ID="LabelTitle" runat="server"></asp:Label>
            </td>
        </tr>
		<tr>
            <td align="left">
                <asp:Label ID="LabelUserName" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <td align="center">
                <asp:Image ID="ImageNow" runat="server" />
            </td>
        </tr>
        <tr>
            <td align="left">
                <asp:Literal ID="LiteralDescription" runat="server"></asp:Literal></td>
        </tr>
    </table>
</div>
<div id="notpublic" runat="server">
    <table class="grid">
        <tr>
            <td align="center">
                <b>本作品被使用者關閉，暫時不公開</b> </td>
        </tr>
    </table>
</div>
<div id="notapprove" runat="server">
    <table class="grid">
        <tr>
            <td align="center">
                <b>本作品被管理員關閉，暫時不公開</b> </td>
        </tr>
    </table>
</div>
<hr />

<%@ Control Language="C#" AutoEventWireup="true" CodeFile="UserInfo.ascx.cs" Inherits="UserControls_UserInfo" %>
<div>
    <table width="100%" class="userinfo">
        <tr>
            <td rowspan="4" style="width: 120px; height: 160px; text-align: center; vertical-align: middle; cursor: pointer;">
                <asp:Image ID="AvatarImage" runat="server" />
            </td>
            <th>
                作者帳號
            </th>
            <td>
                <asp:Label ID="OwnerId" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <th>
                作者名稱
            </th>
            <td>
                <asp:Label ID="DisplayName" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <th>
                作物名稱
            </th>
            <td>
                <asp:Label ID="TextTopic" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <th>
                作品介紹
            </th>
            <td>
                <asp:Literal ID="Description" runat="server"></asp:Literal></td>
        </tr>
    </table>
</div>

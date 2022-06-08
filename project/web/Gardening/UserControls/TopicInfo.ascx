<%@ Control Language="C#" AutoEventWireup="true" CodeFile="TopicInfo.ascx.cs" Inherits="UserControls_TopicInfo" %>
<div>
    <table width="100%" class="userinfo">
        <tr>
            <td rowspan="2" style="width: 120px; height: 80px; text-align: center; vertical-align: middle; cursor: pointer;">
                <asp:Image ID="AvatarImage" runat="server" />
            </td>
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

<%@ Control Language="C#" AutoEventWireup="true" CodeFile="path.ascx.cs" Inherits="path" %>
<div style="padding-top: 5px"></div>
<div class="path" style="float: left; padding-top: 10px">
    目前位置：
</div>
<asp:HiddenField runat="server" ID="HiddenPath" />
<div style="float:left;width:80%">
    <asp:Label runat="server" ID="labPath"></asp:Label>
</div>

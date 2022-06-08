<%@ Page Title="" Language="C#" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true" CodeFile="History_List.aspx.cs" Inherits="Century_History_List" %>

<%@ Register Src="TabText.ascx" TagName="TabMenu" TagPrefix="uc1" %>
<%@ Register Src="path.ascx" TagName="path" TagPrefix="uc2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <!--中間內容區-->
    <!--  目前位置  -->
    <uc2:path runat="server" ID="path1" />
    <%--<div class="path" style="float: left; padding-top: 10px">
        目前位置：
    </div>
    <div style="float: left; width: 90%">
        <ul id="path_menu">
            <li><a href="../mp.asp?mp=1">首頁</a></li>
            <li style='top: 10px;'>></li>
            <li><a href="Events_List.aspx">農業百年發展史</a></li>
        </ul>
    </div>--%>
    <!--標籤區-->
    <uc1:TabMenu runat="server" id="TabMenu1" />
    <asp:Label runat="server" ID="labList"></asp:Label>
    <%=strError %>
</asp:Content>
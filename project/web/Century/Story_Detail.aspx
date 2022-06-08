<%@ Page Title="" Language="C#" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true"
    CodeFile="Story_Detail.aspx.cs" Inherits="Story_Detail" %>

<%@ Register Src="TabText.ascx" TagName="TabMenu" TagPrefix="uc1" %>
<%@ Register Src="path.ascx" TagName="path" TagPrefix="uc2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <!--調整字級    Start-->
    <script type="text/javascript" src="../js/dozoom.js"></script>
    <!--調整字級     End -->
    <!--  目前位置  -->
    <uc2:path runat="server" ID="path1" />
    <!--標籤區-->
    <uc1:TabMenu runat="server" ID="TabMenu1" />
    <!-- 頁面功能 Start -->
    <ul class="Function2">
        <li><a href="#" onclick="window.open('knowledge_cpPrint.aspx<%=Server.UrlEncode(Request.Url.Query).Replace("%3f","?").Replace("%3d","=").Replace("%26","&") %>')"
            class="Print">友善列印</a></li>
        <li>
            <asp:Label ID="labList" runat="server" Text="<a href='Story_List.aspx' class='List'>回列表頁</a>"></asp:Label>
        </li>
        <li>
            <asp:Label ID="labBack" runat="server" Text=""></asp:Label>
        </li>
        <li>
            <asp:Label ID="labNext" runat="server" Text=""></asp:Label>
        </li>
    </ul>
    <!--調整字級    Start-->
    <div style="float: right; margin-left: 18px">
        調整字級： <a onclick="changeFontSize('article','idx1');">
            <img id="idx1" alt="" src="../subject/images/fontsize_1_off.gif" align="absmiddle"
                border="0" name="idx1" style="padding-bottom: 2px;" />
        </a><a onclick="changeFontSize('article','idx2');">
            <img id="idx2" alt="" src="../subject/images/fontsize_2_off.gif" align="absmiddle"
                border="0" name="idx2" style="padding-bottom: 2px;" />
        </a><a onclick="changeFontSize('article','idx3');">
            <img id="idx3" alt="" src="../subject/images/fontsize_3_off.gif" align="absmiddle"
                border="0" name="idx3" style="padding-bottom: 2px;" />
        </a><a onclick="changeFontSize('article','idx4');">
            <img id="idx4" alt="" src="../subject/images/fontsize_4_off.gif" align="absmiddle"
                border="0" name="idx4" style="padding-bottom: 2px;" />
        </a>
    </div>
    <!--調整字級     End -->
    <!-- 內容區 Start -->
    <div id="cp" style="clear: both;">
        <span id="article" name="article">
            <!-- 內容區 Start -->
            <!-- 標題 -->
            <h4>
                <asp:Label ID="labTitle" runat="server" Text=""></asp:Label></h4>
            <!-- 內文 -->
            <asp:Label ID="labContent" runat="server" Text=""></asp:Label>
        </span>
    </div>
</asp:Content>

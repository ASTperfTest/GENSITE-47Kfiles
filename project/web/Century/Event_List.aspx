<%@ Page Title="" Language="C#" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true"
    CodeFile="Event_List.aspx.cs" Inherits="Century_Events_Events_List" %>

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
    <uc1:TabMenu runat="server" ID="TabMenu1" />
    <!--月份-->
    <div id="divMonth">
        <ul class="category" style="padding: 0px 0px 0px 0px;">
            <%
                int nowMonth = DateTime.Now.Month; // 當月份
                string tagMonth = Request.QueryString["month"] ?? "0";
                for (int i = 1; i <= 12; i++)
                {
                    // 指定 Month
                    if (Convert.ToInt32(tagMonth) != 0)
                    {
                        if ((i == Convert.ToInt32(tagMonth)))
                        {
                            Response.Write("<li><a  href='Event_List.aspx?month=" + i + "'><span style='font-weight:bold; Color:#FFFF00'>" + i + "月</span></a></li>");
                        }
                        else
                        {
                            Response.Write("<li><a  href='Event_List.aspx?month=" + i + "'>" + i + "月</a></li>");
                        }
                    }
                    // 不指定 Month
                    else
                    {
                        if ((i == nowMonth))
                        {
                            Response.Write("<li><a  href='Event_List.aspx?month=" + i + "'><span style='font-weight:bold; Color:#FFFF00'>" + i + "月</span></a></li>");
                        }
                        else
                        {
                            Response.Write("<li><a  href='Event_List.aspx?month=" + i + "'>" + i + "月</a></li>");
                        }
                    }
                } 
            %>
        </ul>
    </div>
    <br />
    <br />
    <br />
    <br />
    <!--文章列表-->
    <div id="Event">
        <asp:Label runat="server" ID="labList"></asp:Label>
    </div>
</asp:Content>

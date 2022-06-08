<%@ Page Title="農業知識入口網-好文推薦" Language="C#" MasterPageFile="~/3RSideMasterPage.master"
    AutoEventWireup="true" CodeFile="Recommand_Mylist.aspx.cs" Inherits="Recommand_Mylist" %>

<%@ Register Src="~/knowledge/LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="~/knowledge/TabText.ascx" TagName="TabText" TagPrefix="uc2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <!--中間內容區-->
    <td class="center">
        <!--  目前位置  -->
        <uc2:TabText ID="TabText1" runat="server" />
    <div class="path" style="float: left; padding-top: 10px">
        目前位置：
    </div>
    <div style="float: left; width: 90%">
        <ul id="path_menu">
            <li><a href="../mp.asp?mp=1">首頁</a></li>
            <li style='top: 10px;'>></li>
            <li><a href="Recommand_MyList.aspx">我推薦的文章</a></li>
        </ul>
    </div>
            <h3>
                我推薦的文章</h3>
        <!-- 我推薦的文章 TabList  Start -->
        <div id="MagTabs">
            <asp:label runat="server" id="TabText2" text=""></asp:label>
            <asp:label runat="server" id="labTabtext" text=""></asp:label>
        </div>
        <!-- 我推薦的文章 TabList   End  -->
        <div class="Event">
            <!-- 分頁列表  Start -->
            <div class="Page">
                第<asp:label id="PageNumberText" runat="server" text="" cssclass="Number" />/<asp:label
                    id="TotalPageText" runat="server" text="" cssclass="Number" />頁， 共<asp:label id="TotalRecordText"
                        runat="server" text="" cssclass="Number" />筆資料，
                <asp:linkbutton id="PreviousLink" runat="server" onclick="PreviousLink_Click" visible="false">
        	        <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg" Visible="false" AlternateText="上一頁"></asp:image>        	            
        	        <asp:Label ID="PreviousText" runat="server">上一頁</asp:Label>            
        	    </asp:linkbutton>
                &nbsp; 到第
                <asp:dropdownlist id="PageNumberDDL" autopostback="true" runat="server" />
                頁 &nbsp;
                <asp:linkbutton id="NextLink" runat="server" onclick="NextLink_Click" visible="false">
                    <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg" Visible="false" AlternateText="下一頁"></asp:image>                     	
                    <asp:Label ID="NextText" runat="server" >下一頁</asp:Label>
                </asp:linkbutton>
                &nbsp;，每頁顯示
                <asp:dropdownlist id="PageSizeDDL" autopostback="true" runat="server">
                    <asp:ListItem Selected="True" Value="10">10</asp:ListItem>
                    <asp:ListItem Value="20">20</asp:ListItem>
                    <asp:ListItem Value="30">30</asp:ListItem>
                    <asp:ListItem Value="50">50</asp:ListItem>
                </asp:dropdownlist>
                筆
            </div>
            <!-- 分頁列表   End  -->
            <!-- 文章列表  Start -->
            <asp:label id="TableText" runat="server" text="" />
            <div class="lp">
                <asp:repeater runat="server" id="rptList" onitemdatabound="rptList_ItemDataBound">
                        <HeaderTemplate>
                            <table width="100%">
                                <tr>
                                    <th style="width:70px">推薦日期</th>
                                    <th align="left"  style="width:120px">文章標題</th>
                                    <th align="left">推薦原因</th>
                                    <th style="width:80px" align="left">資料出處</th>
                                    <th style="width:80px" align="left">審查狀態</th>
                                </tr>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <tr>
                                <td><asp:Label runat="server" id="labDate" Text='<%# Eval("aEditDate") %>'></asp:Label></td>
                                <td><asp:HyperLink runat="server" id="hypTitle" NavigateUrl='<%# Eval("URL") %>' Text='<%# Eval("Title") %>' Target="_blank"></asp:HyperLink></td>                                
                                <td><%# Eval("aContent") %></td>
                                <td><%# Eval("Source") %></td>
                                <td><asp:Label runat="server" id="labExam" Text='<%# Eval("Status") %>'></asp:Label></td>
                            </tr>
                        </ItemTemplate> 
                        <FooterTemplate>
                        </table>                       
                        </FooterTemplate>
                    </asp:repeater>
            </div>
            <!-- 文章列表   End  -->
            <div class="gotop">
                <a href="#top">
                    <img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div>
        </div>
    </td>
</asp:Content>

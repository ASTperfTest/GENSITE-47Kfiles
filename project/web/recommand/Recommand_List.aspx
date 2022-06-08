<%@ Page Title="農業知識入口網-好文推薦" Language="C#" MasterPageFile="~/4SideMasterPage.master"
    AutoEventWireup="true" CodeFile="Recommand_List.aspx.cs" Inherits="Recommand_List" %>

<%@ Register Src="~/knowledge/LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
<script type="text/javascript">
    function resetKW() {
        var kw = document.getElementById('<%= txtSearch.ClientID.ToString()  %>');
        kw.value='';
    }

    function DealRecommand() {
        if ("<%=memId%>" == "none") {
            alert("連線逾時或尚未登入，請登入會員");
        }
        else {
            window.location="recommand_Add.aspx";
        }  
    }    
</script>

    <!--中間內容區-->
    <!--  目前位置  -->
    <div class="path" style="float: left; padding-top: 10px">
        目前位置：
    </div>
    <div style="float: left; width: 80%">
        <!--<a href="/recommand/recommand_Add.aspx">-->
        <a>
            <img STYLE="position:absolute; RIGHT:250px; cursor:pointer;" border="0" onclick="DealRecommand()" src="../xslGip/style3/images/推薦bt.gif" alt="推薦好文" />
        </a> 
        <ul id="path_menu">
            <li><a href="../mp.asp?mp=1">首頁</a></li>
            <li style='top: 10px;'>></li>
            <li><a href="Recommand_List.aspx">好文推薦</a></li>
        </ul>        
    </div>
    <div class="pantology" style="clear:both">
        <div class="head"></div>
        <div class="body">
            <h2>好文推薦</h2>
            <div  class="sublayout">
            <asp:Image id="Image1" runat="server" style="border:#aaa 1px solid"></asp:Image>                
                <h4>
                    <asp:label id="lbtitle" runat="server"></asp:label>
                </h4>
                <p>
                <asp:Label id="lbabstract" runat="server"></asp:Label>
                    ...<a href="Recommandinfodetail.aspx">詳全文</a></p><br />
            </div>
        </div>
        <div class="body">
            <!-- 標籤查詢  Start -->
            <h2></h2>
            <div class="browseby">
                <table align="left" style="min-width: 450px;margin-top:10px">
                    <tr>
                        <th valign="middle" align="right" style="width: 100px;">
                            查詢文章：
                        </th>
                        <td align="left">
                            <asp:textbox runat="server" id="txtSearch"></asp:textbox>
                            <asp:button runat="server" id="btnSearch" text="查詢" />
                            <input type="button" value="清除" onclick="resetKW()" />
                        </td>
                    </tr>
                    <tr>
                        <th align="right" style="width: 100px;">
                            分類過濾：
                        </th>
                        <td>
                            <asp:panel runat="server" id="panelTAGs" cssclass="category"></asp:panel>
                        </td>
                    </tr>
                </table>
            </div>
            <!-- 標籤查詢   End  -->
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
            <div class="lp">
                <asp:repeater runat="server" id="rptList">
                        <HeaderTemplate>
                            <table width="100%" cellpadding="0" cellspacing="0">
                                <tr>
                                    <th style="width:10px">&nbsp;</th>
                                    <th align="left" style="width:120px">文章標題</th>
                                    <th align="left">推薦原因</th>
                                </tr>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <tr>
                                <td style="width:10px"><%# Container.ItemIndex +1 %></td>
                                <td align="left" style="width:120px; padding:0 0 0 0;"><asp:HyperLink runat="server" id="hypTitle" NavigateUrl='<%# Eval("URL") %>' Text='<%# Eval("Title") %>' Target="_blank"></asp:HyperLink></td>
                                <td style="line-height:18px"><%# Eval("aContent") %>
                                    <div style="color:Maroon; margin:6px 0 6px;">
                                        推薦會員：<%# GetShowName( Eval("nickname"), Eval("realname"))%>, &nbsp;資料出處：<%# Eval("Source")%>                                        
                                    </div>
                                </td>
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
        <div class="foot">
        </div>
    </div>
</asp:Content>

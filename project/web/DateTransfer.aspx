<%@ Page Language="vb" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="false" 
    CodeFile="DateTransfer.aspx.vb" Inherits="DateTransfer" %>


<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <script src="js/lunar.js" type="text/javascript"></script>
    <link rel="stylesheet" type="text/css" href="./xslgip/style3/css/4seasons.css" />
    <div class="pantology">
            <a title="網頁內容資料" class="Accesskey" accesskey="C" href="#" xmlns:hyweb="urn:gip-hyweb-com">
                :::</a>
                目前位置：<a title="首頁" href="/mp.asp?mp=1">首頁</a>&gt;<a title="首頁" href="#">西曆、農曆換算</a>
        <div class="head"></div>
        <div class="body">
            <h2 style="color:#666600"><b>西曆、農曆換算</b></h2> 
                <table width="100%" border="0" cellpadding="0" cellspacing="1" style="background-color: #CCCCCC">
                    <tr style="background-color: #FFFFFF"><td align="center">
                    <table width="350px">
                        <tr align="left">
                        <td style="width:100px"><asp:RadioButton ID="rdoCountry" runat="server" GroupName="kind" Text="國曆轉農曆" Checked="true" OnClick="monthDays(this)" AutoPostBack="True" /></td>
                        <td rowspan="2">
                            <asp:Label ID="Label1" runat="server" Text="西元"></asp:Label>
                            <asp:TextBox ID="txtYear" runat="server" Width="50px" onblur="leapMonth(this)" onkeypress="isnum()"
                                MaxLength="4">2010</asp:TextBox>
                            <asp:Label ID="Label2" runat="server" Text="年"></asp:Label>
                            <asp:DropDownList ID="ddlMonth" runat="server" onchange="monthDays(this)" 
                                AutoPostBack="True">
                                <asp:ListItem>1</asp:ListItem>
                                <asp:ListItem>2</asp:ListItem>
                                <asp:ListItem>3</asp:ListItem>
                                <asp:ListItem>4</asp:ListItem>
                                <asp:ListItem>5</asp:ListItem>
                                <asp:ListItem>6</asp:ListItem>
                                <asp:ListItem>7</asp:ListItem>
                                <asp:ListItem>8</asp:ListItem>
                                <asp:ListItem>9</asp:ListItem>
                                <asp:ListItem>10</asp:ListItem>
                                <asp:ListItem Value="11"></asp:ListItem>
                                <asp:ListItem>12</asp:ListItem>
                            </asp:DropDownList>
                            <asp:Label ID="Label3" runat="server" Text="月"></asp:Label>
                            <asp:DropDownList ID="ddlDay" runat="server">
                                <asp:ListItem>1</asp:ListItem>
                            </asp:DropDownList>
                            <asp:Label ID="Label4" runat="server" Text="日"></asp:Label>
                        </td>
                        </tr>
                        <tr align="left">
                        <td><asp:RadioButton ID="rdoLunar" runat="server" GroupName="kind" Text="農曆轉國曆" OnClick="monthDays(this)" AutoPostBack="True" /></td><td></td></tr>
                        <tr><td></td></tr>       
                        <tr><td colspan="2"><asp:Button ID="btnSearch" runat="server" Text="查詢" OnClientClick="Submit()" /></td></tr>
                    </table>            
                    </td></tr>
                </table>  
                <div class="jigsawnew">
                    <h5>查詢結果：</h5>
                    <br />
                    <table width="100%"><tr><td align="center">
                        <asp:TextBox ID="txtResult" runat="server" Width="347px" style="BORDER-RIGHT: #FFFFFF 1px solid; BORDER-TOP: #FFFFFF 1px solid; FONT-SIZE: 12px; BORDER-LEFT: #FFFFFF 1px solid; COLOR: #000000; BORDER-BOTTOM: #FFFFFF 1px solid; FONT-FAMILY: Verdana; TEXT-DECORATION: none"></asp:TextBox>
                    </td></tr></table>                    
                </div>
                <asp:HiddenField ID="hidDays" runat="server" />  
        </div>
        <div class="foot"></div>
    </div>
</asp:Content>
<%@ Page Language="C#" AutoEventWireup="true" CodeFile="FarmProduceMap.aspx.cs" Inherits="FarmProduceMap" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>農漁地圖討論管理</title>
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <link type="text/css" rel="stylesheet" href="../css/setstyle.css" />
	<script type="text/javascript" src="http://ajax.microsoft.com/ajax/jquery/jquery-1.4.2.min.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div id="FuncName">
            <h1>
                農漁地圖討論管理<font size="2">【留言清單】</font></h1>
            <div id="ClearFloat"> </div>
        </div>
        <div class="browseby">
            類型：
            <asp:DropDownList ID="type_DropDownList" runat="server" onSelectedIndexChanged="FillUnitDropDowmList" AutoPostBack="True">
                <asp:ListItem Value="">請選擇</asp:ListItem>
                <asp:ListItem Value="0">作物</asp:ListItem>
                <asp:ListItem Value="1">魚種</asp:ListItem>
            </asp:DropDownList>
            &nbsp;
            作物/魚種：
            <asp:DropDownList ID="unit_DropDownList" runat="server" DataTextField="name" DataValueField="name"  
                AppendDataBoundItems="True" Width="200px" onSelectedIndexChanged="ChangeUnit" AutoPostBack="True"> 
                <asp:ListItem Value="">請選擇</asp:ListItem> 
            </asp:DropDownList><br/><br/>
            作者：<asp:TextBox ID="author" runat="server"></asp:TextBox>&nbsp;
            內容：<asp:TextBox ID="content" runat="server"></asp:TextBox>&nbsp;
            作物名稱/魚種名稱：<asp:TextBox ID="unit" runat="server"></asp:TextBox>&nbsp;
            &nbsp;<asp:Button ID="query_btn" runat="server" Text="查詢" OnClick="Query_btn_Click" />
        </div>
        <div id="Page">
            <asp:LinkButton ID="preLink" runat="server" Visible="false" OnClick="preLinkAct">
                <asp:Image ID="preImg" runat="server" ImageUrl="/images/arrow_previous.gif" Visible="false" AlternateText="上一頁" />
                <asp:Label ID="preText" runat="server" Text="上一頁" />
            </asp:LinkButton>&nbsp;
            共&nbsp;<asp:Label ID="totalRecord" runat="server" style="color:#F06000;" />&nbsp;筆資料，
            目前在第<em style="color:#F06000;"><asp:Label ID="pageNum" runat="server" />/<asp:Label ID="totalPage" runat="server" /></em>頁，
            每頁顯示
            <asp:DropDownList ID="pagesize" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ChangePageSize">
                <asp:ListItem Value="15">15</asp:ListItem>
                <asp:ListItem Value="30">30</asp:ListItem>
                <asp:ListItem Value="50">50</asp:ListItem>
            </asp:DropDownList>
            筆，跳至第
            <asp:DropDownList ID="nowpage" runat="server" 
                AutoPostBack="True" OnSelectedIndexChanged="ChangePage">
            </asp:DropDownList>
            頁&nbsp;
            <asp:LinkButton ID="nextLink" runat="server" Visible="false" OnClick="nextLinkAct">
                <asp:Image ID="nextImg" runat="server" ImageUrl="/images/arrow_next.gif" Visible="false" AlternateText="下一頁" />
                <asp:Label ID="nextText" runat="server" Text="下一頁" />
            </asp:LinkButton>
        </div>
        <asp:Repeater ID="rpList" runat="server">
            <HeaderTemplate>
                <table cellspacing="0" id="ListTable">
                    <tr>
                        <th class="eTableLable" width="2%">
                            <!--<input type="button" value="全選" class="cbutton" name="checkAllBtn" onclick="CheckAll()"/>-->
                        </th>
                        <th class="eTableLable" width="3%">
                            預覽
                        </th>
                        <th class="eTableLable" width="3%">
                            編輯
                        </th>
                        <th class="eTableLable" width="6%">
                            作物/魚種
                        </th>
                        <th class="eTableLable" width="10%">
                            作者
                        </th>
					    <th class="eTableLable" style="width:12%">
                            日期
                        </th>
                        <th class="eTableLable" width="6%">
                            公開/不公開
                        <th class="eTableLable">
                            內容
                        </th>
                    </tr>
            </HeaderTemplate>
            <ItemTemplate>
                <tr>
                    <td class="eTableContent">
                        <asp:CheckBox ID="CheckBox1" runat="server" ToolTip='<%#Eval("icuitem")%>'/>
                    </td>
                    <td class="eTableContent">
                        <a target="_blank" href="<%# StrFunc.GetWebSetting("WebURL", "appSettings").Replace("\\", "/") %>/jigsaw2010/detail.aspx?item=<%#Eval("unitId")%>">
                            view
                        </a>
                    </td>
                    <td class="eTableContent">
                        <asp:Button ID="Button1" runat="server" Text="編輯" OnClick="Edit" ToolTip='<%#Eval("icuitem")%>'/>
                    </td>
                    <td class="eTableContent">
                        <%#Eval("sTitle")%>
                    </td>
                    <td class="eTableContent">
                        <%#Eval("iEditor")%>
                    </td>
                    <td class="eTableContent">
                        <%#Eval("Created_Date")%>
                    </td>
                    <td class="eTableContent">
                        <%#FixYN(Eval("fCTUPublic"))%>
                    </td>
                    <td class="eTableContent">
                    <div>
                        <asp:Label ID="Label1" runat="server" Text='<%#Eval("xBody")%>'></asp:Label>
                        <asp:TextBox ID="TextBox1" runat="server" Width="90%" style="display:none;"></asp:TextBox>
                        <asp:Button ID="save" runat="server" style="display:none;" ToolTip='<%#Eval("icuitem")%>' OnClick="Save" Text="儲存" />
                    </div>
                    </td>
               </tr>
           </ItemTemplate>
           <FooterTemplate>
               </table>
           </FooterTemplate>
        </asp:Repeater>
        <div align="center">
            <asp:Button ID="public" runat="server" class="cbutton" OnClick="DealPublic" Text="公開" />
            <asp:Button ID="nopublic" runat="server" class="cbutton" OnClick="DealPublic" Text="不公開" />
        </div>
    </form>
</body>
</html>

<script type="text/javascript" language="javascript">

    
</script>


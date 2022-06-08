<%@ Page Language="C#" AutoEventWireup="true" CodeFile="FarmProduceMap.aspx.cs" Inherits="FarmProduceMap" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>�A���a�ϰQ�׺޲z</title>
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <link type="text/css" rel="stylesheet" href="../css/setstyle.css" />
	<script type="text/javascript" src="http://ajax.microsoft.com/ajax/jquery/jquery-1.4.2.min.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div id="FuncName">
            <h1>
                �A���a�ϰQ�׺޲z<font size="2">�i�d���M��j</font></h1>
            <div id="ClearFloat"> </div>
        </div>
        <div class="browseby">
            �����G
            <asp:DropDownList ID="type_DropDownList" runat="server" onSelectedIndexChanged="FillUnitDropDowmList" AutoPostBack="True">
                <asp:ListItem Value="">�п��</asp:ListItem>
                <asp:ListItem Value="0">�@��</asp:ListItem>
                <asp:ListItem Value="1">����</asp:ListItem>
            </asp:DropDownList>
            &nbsp;
            �@��/���ءG
            <asp:DropDownList ID="unit_DropDownList" runat="server" DataTextField="name" DataValueField="name"  
                AppendDataBoundItems="True" Width="200px" onSelectedIndexChanged="ChangeUnit" AutoPostBack="True"> 
                <asp:ListItem Value="">�п��</asp:ListItem> 
            </asp:DropDownList><br/><br/>
            �@�̡G<asp:TextBox ID="author" runat="server"></asp:TextBox>&nbsp;
            ���e�G<asp:TextBox ID="content" runat="server"></asp:TextBox>&nbsp;
            �@���W��/���ئW�١G<asp:TextBox ID="unit" runat="server"></asp:TextBox>&nbsp;
            &nbsp;<asp:Button ID="query_btn" runat="server" Text="�d��" OnClick="Query_btn_Click" />
        </div>
        <div id="Page">
            <asp:LinkButton ID="preLink" runat="server" Visible="false" OnClick="preLinkAct">
                <asp:Image ID="preImg" runat="server" ImageUrl="/images/arrow_previous.gif" Visible="false" AlternateText="�W�@��" />
                <asp:Label ID="preText" runat="server" Text="�W�@��" />
            </asp:LinkButton>&nbsp;
            �@&nbsp;<asp:Label ID="totalRecord" runat="server" style="color:#F06000;" />&nbsp;����ơA
            �ثe�b��<em style="color:#F06000;"><asp:Label ID="pageNum" runat="server" />/<asp:Label ID="totalPage" runat="server" /></em>���A
            �C�����
            <asp:DropDownList ID="pagesize" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ChangePageSize">
                <asp:ListItem Value="15">15</asp:ListItem>
                <asp:ListItem Value="30">30</asp:ListItem>
                <asp:ListItem Value="50">50</asp:ListItem>
            </asp:DropDownList>
            ���A���ܲ�
            <asp:DropDownList ID="nowpage" runat="server" 
                AutoPostBack="True" OnSelectedIndexChanged="ChangePage">
            </asp:DropDownList>
            ��&nbsp;
            <asp:LinkButton ID="nextLink" runat="server" Visible="false" OnClick="nextLinkAct">
                <asp:Image ID="nextImg" runat="server" ImageUrl="/images/arrow_next.gif" Visible="false" AlternateText="�U�@��" />
                <asp:Label ID="nextText" runat="server" Text="�U�@��" />
            </asp:LinkButton>
        </div>
        <asp:Repeater ID="rpList" runat="server">
            <HeaderTemplate>
                <table cellspacing="0" id="ListTable">
                    <tr>
                        <th class="eTableLable" width="2%">
                            <!--<input type="button" value="����" class="cbutton" name="checkAllBtn" onclick="CheckAll()"/>-->
                        </th>
                        <th class="eTableLable" width="3%">
                            �w��
                        </th>
                        <th class="eTableLable" width="3%">
                            �s��
                        </th>
                        <th class="eTableLable" width="6%">
                            �@��/����
                        </th>
                        <th class="eTableLable" width="10%">
                            �@��
                        </th>
					    <th class="eTableLable" style="width:12%">
                            ���
                        </th>
                        <th class="eTableLable" width="6%">
                            ���}/�����}
                        <th class="eTableLable">
                            ���e
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
                        <asp:Button ID="Button1" runat="server" Text="�s��" OnClick="Edit" ToolTip='<%#Eval("icuitem")%>'/>
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
                        <asp:Button ID="save" runat="server" style="display:none;" ToolTip='<%#Eval("icuitem")%>' OnClick="Save" Text="�x�s" />
                    </div>
                    </td>
               </tr>
           </ItemTemplate>
           <FooterTemplate>
               </table>
           </FooterTemplate>
        </asp:Repeater>
        <div align="center">
            <asp:Button ID="public" runat="server" class="cbutton" OnClick="DealPublic" Text="���}" />
            <asp:Button ID="nopublic" runat="server" class="cbutton" OnClick="DealPublic" Text="�����}" />
        </div>
    </form>
</body>
</html>

<script type="text/javascript" language="javascript">

    
</script>


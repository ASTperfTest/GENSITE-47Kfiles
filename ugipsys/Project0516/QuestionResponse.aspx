<%@ Page Language="VB" AutoEventWireup="false" CodeFile="QuestionResponse.aspx.vb" Inherits="QuestionResponse" EnableEventValidation="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link href="../css/list.css" rel="stylesheet" type="text/css">
    <link href="../css/layout.css" rel="stylesheet" type="text/css">
    <title>資料管理／問題單回應</title>
</head>
<body>
    <form id="form1" runat="server">
        <div id="FuncName">
            <h1>
                資料管理／問題單回應</h1>            
            <asp:Panel ID="PanelSearch" runat="server">
            <br/>
            <hr size="1" />
                <table border="0" cellspacing="1" cellpadding="2" style="font-size:12px">
                    <tr>
                        <td style="background-color: #d0f1bd; width: 100px; text-align: right">問題來源</td>
                        <td class="eTableContent">
                            <asp:DropDownList ID="searchSource" runat="server">
                                <asp:ListItem Value="">全部</asp:ListItem>
                                <asp:ListItem Value="1">問題反應-知識家</asp:ListItem>
                                <asp:ListItem Value="2">討論檢舉-知識家</asp:ListItem>
                                <asp:ListItem Value="3">問題反應-知識庫</asp:ListItem>
                                <asp:ListItem Value="4">問題反應-主題館</asp:ListItem>
                                <asp:ListItem Value="99">問題反應-最新消息</asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #d0f1bd; width: 100px; text-align: right">處理狀態</td>
                        <td class="eTableContent">
                            <asp:DropDownList ID="searchSTATUS" runat="server">
                                <asp:ListItem Value="">全部</asp:ListItem>
                                <asp:ListItem Value="0" Selected="true">尚未處理</asp:ListItem>
                                <asp:ListItem Value="1">處理完畢</asp:ListItem>
                                <asp:ListItem Value="9">刪除</asp:ListItem>
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #d0f1bd; width: 100px; text-align: right">問題反應人</td>
                        <td class="eTableContent">
                            <asp:TextBox ID="searchSender" runat="server" Width="300" />(帳號、真實姓名、暱稱)
                        </td>
                    </tr>
                    <tr>
                        <td style="background-color: #d0f1bd; width: 100px; text-align: right">反應內容</td>
                        <td class="eTableContent">
                            <asp:TextBox ID="searchDescription" runat="server"  Width="300" />
                            <asp:Button ID="searchSubmit" runat = "server" Text="送出" />
                        </td>
                    </tr>
                </table>
                <hr size="1" />
            </asp:Panel>
            <div id="Page">
                <table border="0" runat="server" id="PageTable">
                    <tr>
                        <td>
                            <asp:Label ID="PageInfo" runat="server" Text="共<em>{0}</em>筆資料，目前在第<em>{1}/{2}</em>頁每頁顯示" /></td>
                        <td>
                            <asp:DropDownList ID="PageDDSize" runat="server" AutoPostBack="true" OnSelectedIndexChanged="OnSetNewPagesize">
                                <asp:ListItem Text="15" Value="15" />
                                <asp:ListItem Text="30" Value="30" />
                                <asp:ListItem Text="50" Value="50" />
                            </asp:DropDownList>
                        </td>
                        <td width="30" align="left">筆</td>
                        <td>
                            <img src="/images/arrow_previous.gif" alt="上一頁"></td>
                        <td>
                            <asp:LinkButton ID="PageBack" runat="server" OnClick="OnClick_Back">上一頁</asp:LinkButton></td>
                        <td>
                            <asp:DropDownList ID="PageDDList" runat="server" AutoPostBack="true" OnSelectedIndexChanged="OnGoNewPage" />
                        </td>
                        <td>
                            <asp:LinkButton ID="PageNext" runat="server" OnClick="OnClick_Next">下一頁</asp:LinkButton></td>
                        <td>
                            <img src="/images/arrow_next.gif" alt="下一頁"></td>
                    </tr>
                </table>
            </div>
            <div>
                <asp:Panel ID="PanelQueryResult" runat="server">
                    <asp:GridView ID="GridView3" runat="server" AutoGenerateColumns="False" BorderStyle="None" Width="100%" CellPadding="4" DataKeyNames="SEQ" GridLines="Horizontal" OnPageIndexChanging="GridView1_PageIndexChanging" OnRowEditing="GridView1_PageIndexChanging_RowEditing" ForeColor="#333333" HeaderStyle-Height="30">
                        <RowStyle BackColor="Transparent" ForeColor="#333333" Font-Size="9pt" BorderStyle="Dotted"/>
                        <Columns>
                            <asp:CommandField ShowEditButton="True" ItemStyle-Wrap="false" />
                            <asp:BoundField DataField="SEQ" HeaderText="問題單號" InsertVisible="False" ReadOnly="True">
                                <HeaderStyle HorizontalAlign="Right" Wrap="False" />
                                <ItemStyle HorizontalAlign="Right" />
                            </asp:BoundField>
                            <asp:BoundField DataField="type" HeaderText="問題來源" InsertVisible="False" ReadOnly="True" >
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" Wrap="False" />
                            </asp:BoundField>
                            <asp:TemplateField HeaderText="問題反應人">
                                <ItemTemplate>
                                    <asp:Label ID="Label2" runat="server" Text='<%# Bind("CREATOR") %>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Left" Wrap="False" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="反應日期">
                                <ItemTemplate>
                                    <asp:Label ID="Label3" runat="server" Text='<%# String.Format("{0:yyyy/MM/dd hh:mm:ss}", Eval("CREATION_DATETIME"))  %>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Left" Wrap="False" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>
                            <asp:HyperLinkField DataNavigateUrlFields="SOURCE_URL" Text="View" HeaderText="問題網頁" Target="_blank" SortExpression="SOURCE_URL">
                                <HeaderStyle HorizontalAlign="Left" Wrap="False" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:HyperLinkField>
                            <asp:TemplateField HeaderText="問題內容描述">
                                <ItemTemplate>
                                    <asp:Label ID="labDESCRIPTION" runat="server" Text='<%# Eval("DESCRIPTION") %>'></asp:Label>
                                </ItemTemplate>
                                <HeaderStyle HorizontalAlign="Left" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:TemplateField>                            
                            <asp:BoundField DataField="RESPONSE" HeaderText="處理敘述" SortExpression="RESPONSE">
                                <HeaderStyle HorizontalAlign="Left" Wrap="False" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                            <asp:BoundField DataField="STATUS" HeaderText="處理狀態" SortExpression="STATUS">
                                <HeaderStyle HorizontalAlign="Left" Wrap="False" />
                                <ItemStyle HorizontalAlign="Left" />
                            </asp:BoundField>
                        </Columns>                        
                        <FooterStyle BackColor="#507CD1" ForeColor="White" Font-Bold="True" />
                        <PagerStyle BackColor="#D0F1BD" ForeColor="Black" Font-Size="9pt" HorizontalAlign="Center" />
                        <SelectedRowStyle BackColor="#339966" ForeColor="White" Font-Bold="True" />
                        <HeaderStyle BackColor="#D0F1BD" Font-Bold="False" ForeColor="Black" Font-Size="9pt" BorderStyle="Ridge" BorderColor="#ECE9D8" BorderWidth="3px" Height="30px" />
                        <AlternatingRowStyle BackColor="White" />
                        <EditRowStyle BackColor="#2461BF"/>
                    </asp:GridView>
                </asp:Panel>
                <asp:Panel ID="Panel1" runat="server" Visible="false">
                    <br />
                    <hr size="1" />
                    <asp:Table ID="Table1" runat="server" Width="80%" CellSpacing="2">
                        <asp:TableRow>
                            <asp:TableCell Width="20%" HorizontalAlign="Right">
                                問題單編號
                            </asp:TableCell>
                            <asp:TableCell Width="30%">
                                <asp:TextBox ID="txtSEQ" runat="server" Enabled="false"></asp:TextBox>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell Width="20%" HorizontalAlign="Right">
                                <asp:Label ID="lblRESPONSE" runat="server" Text="處理敘述"></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell Width="80%" ColumnSpan="3">
                                <asp:TextBox ID="textboxRESPONSE" runat="server" TextMode="MultiLine" Width="75%" Rows="5"></asp:TextBox>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell Width="20%" HorizontalAlign="Right">
                                <asp:Label ID="lblSTATUS" runat="server" Text="處理狀態"></asp:Label>
                            </asp:TableCell>
                            <asp:TableCell>
                                <asp:DropDownList ID="drpSTATUS" runat="server">
                                    <asp:ListItem Value="0">尚未處理</asp:ListItem>
                                    <asp:ListItem Value="1">處理完畢</asp:ListItem>
                                    <asp:ListItem Value="9">刪除</asp:ListItem>
                                </asp:DropDownList>
                            </asp:TableCell>
                        </asp:TableRow>
                        <asp:TableRow>
                            <asp:TableCell>
                                <asp:Button ID="buttonSave" runat="server" OnClick="buttonSave_Click" Text="儲存" />
                                <asp:Button ID="buttonClose" runat="server" CausesValidation="false" OnClick="buttonClose_Click" Text="關閉" />
                            </asp:TableCell>
                        </asp:TableRow>
                    </asp:Table>
                </asp:Panel>
            </div>
    </form>
</body>
</html>

<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="kpi_set_scored.aspx.vb" Inherits="kpi_set_scored" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <asp:GridView ID="GridView1" runat="server" CellPadding="4" GridLines="Horizontal" AutoGenerateColumns="False" BackColor="White" BorderColor="#336666" BorderStyle="Double" BorderWidth="3px">
        <FooterStyle BackColor="White" ForeColor="#333333" />
        <RowStyle BackColor="White" ForeColor="#333333" />
        <Columns>
            <asp:BoundField DataField="名稱" HeaderText="名稱" />
            <asp:HyperLinkField DataNavigateUrlFields="Rank0_4" DataNavigateUrlFormatString="kpi_set_score.aspx?id={0}"
                Text="設定" />
        </Columns>
        <PagerStyle BackColor="#336666" ForeColor="White" HorizontalAlign="Center" />
        <SelectedRowStyle BackColor="#339966" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#336666" Font-Bold="True" ForeColor="White" />
    </asp:GridView>
    <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/kpi_set_index.aspx">上一頁</asp:HyperLink>














</asp:Content>


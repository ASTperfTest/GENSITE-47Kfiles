<%@ Control Language="C#" AutoEventWireup="true" CodeFile="VoteRecordList.ascx.cs" Inherits="UserControls_VoteRecordList" %>
<div>
    <table width="100%">
        <tr>
            <td align="right">
                <asp:Button ID="HideRecord" runat="server" Text="關閉" OnClick="HideRecord_Click" />
            </td>
        </tr>
        <tr>
            <td align="center">
                <%= OwnerId %>
                的投票記錄 </td>
        </tr>
        <tr>
            <td>
                <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" Width="100%" BackColor="White" BorderColor="#E7E7FF" BorderStyle="None" BorderWidth="1px" CellPadding="3" GridLines="Horizontal">
                    <RowStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" />
                    <Columns>
                        <asp:BoundField DataField="VoterName" HeaderText="投票者" />
                        <asp:BoundField DataField="VoteDate" HeaderText="投票時間" />
                    </Columns>
                    <FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
                    <PagerStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" HorizontalAlign="Right" />
                    <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="#F7F7F7" />
                    <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#F7F7F7" />
                    <AlternatingRowStyle BackColor="#F7F7F7" />
                    <EmptyDataTemplate>
                        <asp:Table ID="Table1" runat="server" Width="100%">
                            <asp:TableHeaderRow Width="100%" BackColor="#4A3C8C" Font-Bold="True" ForeColor="#F7F7F7">
                                <asp:TableHeaderCell Text="投票者" />
                                <asp:TableHeaderCell Text="投票時間" />
                            </asp:TableHeaderRow>
                            <asp:TableRow>
                                <asp:TableCell ColumnSpan="2">目前無資料</asp:TableCell>
                            </asp:TableRow>
                        </asp:Table>
                    </EmptyDataTemplate>
                </asp:GridView>
            </td>
        </tr>
    </table>
</div>

<%@ Page Language="C#" AutoEventWireup="true" CodeFile="topiclist.aspx.cs" Inherits="topiclist"
    MasterPageFile="Default2.master" %>

<%@ Register Src="UserControls/UserInfo.ascx" TagName="UserInfo" TagPrefix="uc1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cp" runat="Server">
    <asp:HiddenField ID="useManagementMasterPage" runat="server" />
    <div id="content">
        <input type="hidden" name="Action" id="Action" />
        <table width="972px">
            <tr>
                <td>
                    <table>
                        <tr>
                            <td>
                                <b>排序依據：</b>
                            </td>
                            <td class="text">
                                <asp:RadioButtonList ID="RadioButtonList2" runat="server" RepeatDirection="Horizontal">
                                    <asp:ListItem Text="帳號" Value="ID" Selected="True" />
                                    <asp:ListItem Text="作品名稱" Value="Topic" />
                                    <asp:ListItem Text="修改時間" Value="LastModifyDateTime" />
                                    <asp:ListItem Text="日誌數量" Value="EntryCount" />
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <b>順序：</b>
                            </td>
                            <td class="text">
                                <asp:RadioButtonList ID="RadioButtonList1" runat="server" RepeatDirection="Horizontal">
                                    <asp:ListItem Text="遞增" Value="" Selected="True" />
                                    <asp:ListItem Text="遞減" Value="DESC" />
                                </asp:RadioButtonList>
                            </td>
                            <td>
                                <asp:Button ID="Sort" runat="server" Text="更新排序" OnClick="Sort_Click" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td width="80%" align="center" colspan="4">
                    <asp:GridView ID="GridView1" runat="server" Width="100%" AutoGenerateColumns="False"
                        OnRowDataBound="GridView1_RowDataBound" AllowPaging="True" OnPageIndexChanging="GridView1_PageIndexChanging"
                        PageSize="5">
                        <Columns>
                            <asp:TemplateField HeaderText="作品集">
                                <ItemTemplate>
                                    <uc1:UserInfo ID="UserInfo1" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <PagerSettings NextPageText="" />
                        <PagerTemplate>
                            <table class="pageTable">
                                <tr class="pageStyle">
                                    <td class="pager">
                                        <asp:LinkButton ID="LinkButton1" runat="server" CommandName="Page" CommandArgument="First"
                                            Text="|<" />
                                        <asp:LinkButton ID="LinkButton2" runat="server" CommandName="Page" CommandArgument="Prev"
                                            Text="上一頁" />
                                        <span class="curPage">
                                            <asp:Label ID="PageIndexTextBox" runat="server" />
                                        </span>
                                        <asp:LinkButton ID="LinkButton3" runat="server" CommandName="Page" CommandArgument="Next"
                                            Text="下一頁" />
                                        <asp:LinkButton ID="LinkButton4" runat="server" CommandName="Page" CommandArgument="Last"
                                            Text=">|" />
                                    </td>
                                    <td width="30%" align="center" class="pager">
                                        <div class="text">
                                            有 <span class="total_page">
                                                <asp:Label ID="TotalTuplesTextBox" runat="server" />
                                            </span>筆資料 共 <span class="total_page">
                                                <asp:Label ID="TotalPagesTextBox" runat="server" />
                                            </span>頁 每頁
                                            <asp:DropDownList ID="NumOfTuplesDropDownList" runat="server" AutoPostBack="True"
                                                OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged" CssClass="backwhite">
                                                <asp:ListItem>5</asp:ListItem>
                                                <asp:ListItem>10</asp:ListItem>
                                                <asp:ListItem>50</asp:ListItem>
                                            </asp:DropDownList>
                                            筆
                                        </div>
                                    </td>
                                </tr>
                            </table>
                        </PagerTemplate>
                        <EmptyDataTemplate>
                            <asp:Table ID="Table1" runat="server" Width="100%">
                                <asp:TableHeaderRow Width="100%">
                                    <asp:TableHeaderCell Text="參賽者" />
                                    <asp:TableHeaderCell Text="目前票數" />
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

    <script language='javascript' type="text/javascript">
        $(document).ready(function() {
            //Tab Menu
            $("#tabInfo").addClass("btnDistrict");
            $("#tabList").addClass("btnDistrict");
            $("#tabVote").addClass("btnDistrict now");
            $("#tabFinal").addClass("btnDistrict");
        });

        function vote(id) {
            $("#Action").val('Vote|' + id);
        }

        function Export() {
            $("#Action").val('Export');
            document.aspnetForm.submit();
        }

        function CheckVoteRecord(id) {
            $("#Action").val('CheckRecord|' + id);
        }
    </script>

</asp:Content>

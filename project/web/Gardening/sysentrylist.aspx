<%@ Page Title="" Language="C#" MasterPageFile="Default2.master" AutoEventWireup="true"
    CodeFile="sysentrylist.aspx.cs" Inherits="sysentrylist" %>

<%@ Register Src="UserControls/EntryControl.ascx" TagName="EntryControl" TagPrefix="uc1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cp" runat="Server">
    <asp:HiddenField ID="useManagementMasterPage" runat="server" />
    <div id="content">
        <input type="hidden" name="Action" id="Action" value="Non" />
        <table width="1003px" id="List" runat="server">
            <tr>
                <td>
                    <table>
                        <tr>
                            <td>
                                <b>排序依據：</b>
                            </td>
                            <td class="text">
                                <asp:RadioButtonList ID="rdobtnSortBy" runat="server" RepeatDirection="Horizontal">
                                    <asp:ListItem Text="帳號" Value="OWNER_ID"  />
                                    <asp:ListItem Text="拍攝日期" Value="Date"  Selected="True" />
                                    <asp:ListItem Text="上傳日期" Value="CreateDateTime" />											
                                    <asp:ListItem Text="標題" Value="TITLE" />
                                    <asp:ListItem Text="修改時間" Value="LastModifyDateTime"/>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <b>順序：</b>
                            </td>
                            <td class="text">
                                <asp:RadioButtonList ID="rdobtnOrder" runat="server" RepeatDirection="Horizontal">
                                    <asp:ListItem Text="遞增" Value=""  />
                                    <asp:ListItem Text="遞減" Value="DESC" Selected="True"/>
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
                <td align="center">
                    <asp:GridView ID="GridView1" runat="server" Width="100%" AutoGenerateColumns="False"
                        OnRowDataBound="GridView1_RowDataBound" AllowPaging="True" OnPageIndexChanging="GridView1_PageIndexChanging"
                        PageSize="5">
                        <Columns>
                            <asp:TemplateField HeaderText="日誌">
                                <ItemTemplate>
                                    <uc1:entrycontrol id="EntryControl1" runat="server" />
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

        });
        function ApproveEntry(id) {
            $('#Action').val('ApproveEntry|' + id);
            document.forms[0].submit();
        }
        function HideEntry(id) {
            $('#Action').val('HideEntry|' + id);
            document.forms[0].submit();
        }
		function ConfirmDelete(id) {
            if (confirm('是否確認刪除？\r\n注意：經刪除後即無法回復！')) {
                $('#Action').val('DeleteEntry|' + id);
                document.aspnetForm.submit();
            }
        }
        function EditEntry(id) {
            $('#Action').val('EditEntry|' + id);
            document.aspnetForm.submit();
        }
    </script>

</asp:Content>

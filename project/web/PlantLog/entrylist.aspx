<%@ Page Language="C#" AutoEventWireup="true" CodeFile="entrylist.aspx.cs" Inherits="entrylist"
    MasterPageFile="Default.master" %>

<%@ Register Src="UserControls/EntryControl.ascx" TagName="EntryControl" TagPrefix="uc1" %>
<%@ Register Src="UserControls/UserInfo.ascx" TagName="UserInfo" TagPrefix="uc2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cp" runat="Server">
    <asp:HiddenField ID="useManagementMasterPage" runat="server" />
    <div id="content">
        <input type="hidden" name="Action" id="Action" value="Non" />
        <table width="100%" id="List" runat="server">
            <tr>
                <td class="menu">
                    <table>
                        <tr>
                            <td>
                                <a href="javascript:location.href='vote.aspx';">回上一頁</a>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr id="trApproveUserInfo" runat="server">
                <td class="menu">
                    <table>
                        <tr>
                            <td>
                                <a href="javascript:ApproveOwnerInfo();">開啟個人資料</a>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr id="trHideUserInfo" runat="server">
                <td class="menu">
                    <table>
                        <tr>
                            <td>
                                <a href="javascript:HideOwnerInfo();">隱藏個人資料</a>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr id="trUserInfo" runat="server">
                <td>
                    <uc2:UserInfo ID="UserInfo1" runat="server" />
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:Repeater ID="entryRepeater" runat="server" OnItemDataBound="entryRepeater_ItemDataBound">
                        <HeaderTemplate>
                            <hr />
                        </HeaderTemplate>
                        <ItemTemplate>
                            <uc1:EntryControl ID="EntryControl1" runat="server" />
                        </ItemTemplate>
                    </asp:Repeater>
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
        function ApproveEntry(id) {
            $('#Action').val('ApproveEntry|' + id);
            document.forms[0].submit();
        }
        function HideOwnerInfo() {
            $('#Action').val('HideOwnerInfo');
            document.forms[0].submit();
        }
        function ApproveOwnerInfo() {
            $('#Action').val('ApproveOwnerInfo');
            document.forms[0].submit();
        }
        function HideEntry(id) {
            $('#Action').val('HideEntry|' + id);
            document.forms[0].submit();
        }
    </script>

</asp:Content>

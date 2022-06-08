<%@ Page Language="C#" AutoEventWireup="true" CodeFile="uploadentry.aspx.cs" Inherits="uploadentry" MasterPageFile="Default.master" %>

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
                                <a href="javascript:UpdateOwnerInfo();">更新個人資料</a> </td>
                            <td>
                                <a href="javascript:NewLog();">新增日誌</a> </td>
                            <td>
                                <a href="javascript:history.go(-1);">回上一頁</a> </td>    
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
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
            $("#tabList").addClass("btnDistrict now");
            $("#tabVote").addClass("btnDistrict");
            $("#tabFinal").addClass("btnDistrict");
        });
        function ConfirmDelete(id) {
            if (confirm('是否確認刪除？\r\n注意：經刪除後即無法回復！')) {
                $('#Action').val('DeleteEntry|' + id);
                document.forms[0].submit();
            }
        }
        function EditEntry(id) {
            $('#Action').val('EditEntry|' + id);
            document.forms[0].submit();
        }
        function UpdateOwnerInfo() {
            $('#Action').val('UpdateOwnerInfo');
            document.forms[0].submit();
        }
        function NewLog() {
            $('#Action').val('NewLog');
            document.forms[0].submit();
        }
    </script>

</asp:Content>

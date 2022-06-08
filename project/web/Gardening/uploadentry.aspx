<%@ Page Language="C#" AutoEventWireup="true" CodeFile="uploadentry.aspx.cs" Inherits="uploadentry" MasterPageFile="Default2.master" %>

<%@ Register Src="UserControls/EntryControl.ascx" TagName="EntryControl" TagPrefix="uc1" %>
<%@ Register Src="UserControls/UserInfo.ascx" TagName="UserInfo" TagPrefix="uc2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cp" runat="Server">
    <asp:HiddenField ID="useManagementMasterPage" runat="server" />
    <div id="content">
        <input type="hidden" name="Action" id="Action" value="Non" />
        <table width="972px" id="List" runat="server">
            <tr>
                <td class="menu">
                    <table>
                        <tr>
                            <td>
                                <a href="javascript:UpdateTopic();">更新作品資料</a> </td>
                            <td>
                                <a href="javascript:NewLog();">新增日誌</a> </td>
                            <td>
                                <a href="javascript:DeleteTopic();">刪除作品資料</a> </td>
                            <td>
                                <a href="javascript:Cancel();">取消</a> </td>
							<td id="trApproveUserInfo" runat="server">
                                <a href="javascript:ApproveOwnerInfo();">開啟作品</a> </td> 				
							<td id="trHideUserInfo" runat="server">
                                <a href="javascript:HideOwnerInfo();">隱藏作品</a> </td> 					
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
                document.aspnetForm.submit();
            }
        }
        function EditEntry(id) {
            $('#Action').val('EditEntry|' + id);
            document.aspnetForm.submit();
        }
        function UpdateTopic() {
            $('#Action').val('UpdateTopic');
            document.aspnetForm.submit();
        }
        function NewLog() {
            $('#Action').val('NewLog');
            document.aspnetForm.submit();
        }
        function DeleteTopic() {
            if (confirm('是否確認刪除？\r\n注意：經刪除後即無法回復！')) {
                $('#Action').val('DeleteTopic');
                document.aspnetForm.submit();
            }
        }
		function ApproveEntry(id) {
            $('#Action').val('ApproveEntry|' + id);
            document.aspnetForm.submit();
        }
		function HideEntry(id) {
            $('#Action').val('HideEntry|' + id);
            document.aspnetForm.submit();
        }
		function HideOwnerInfo() {
            $('#Action').val('HideOwnerInfo');
            document.aspnetForm.submit();
        }
        function ApproveOwnerInfo() {
            $('#Action').val('ApproveOwnerInfo');
            document.aspnetForm.submit();
        }
		function Cancel() {
            $('#Action').val('Cancel');
            document.aspnetForm.submit();
        }
    </script>

</asp:Content>

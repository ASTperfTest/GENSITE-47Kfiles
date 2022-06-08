<%@ Page Language="C#" ValidateRequest="false" AutoEventWireup="true" CodeFile="editownerinfo.aspx.cs" Inherits="editownerinfo" MasterPageFile="Default.master" %>

<asp:Content ID="Content1" ContentPlaceHolderID="cp" runat="Server">

    <script language="javascript" type="text/javascript" src="js/jquery.autogrow.js"></script>
    <asp:HiddenField ID="useManagementMasterPage" runat="server" />
    <div id="content">
        <table width="100%">
            <tr>
                <td class="menu">
                    <input type="hidden" name="Action" id="Action" value="Non" />
                    <table>
                        <tr>
                            <td id="NewImg" runat="server">
                                <a href="javascript:UpdateAvatar();">更新您的照片</a> </td>
                            <td>
                                <a href="javascript:CheckUpdate();;">儲存</a> </td>
                            <td>
                                <a href="javascript:Cancel();">取消</a> </td>
                            <td>
                                <a href="javascript:Hint();">欄位說明</a> </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td valign="top">
                    <table class="userinfo">
                        <tr>
                            <th>
                                我的帳號
                            </th>
                        </tr>
                        <tr>
                            <td>
                                <asp:Label ID="LabelOwnerId" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                我的姓名
                            </th>
                        </tr>
                        <tr>
                            <td>
                                <asp:Label ID="DisplayName" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                我的植物
                            </th>
                        </tr>
                        <tr>
                            <td>
                                <asp:TextBox ID="NewTopic" runat="server" MaxLength="50" Width="50%"></asp:TextBox>
                                <div class="descriptiveText" style="display: none">
                                    請寫出您要種植的植物正式名稱</div>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                作品說明
                            </th>
                        </tr>
                        <tr>
                            <td>
                                <textarea id="Des" runat="server" cols="60"></textarea>
                                <div class="descriptiveText" style="display: none">
                                    您可以在此區做簡單的自我介紹及作品說明</div>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                您的照片
                            </th>
                        </tr>
                        <tr id="ImgUp" runat="server">
                            <td>
                                <asp:FileUpload ID="ImgUpload" runat="server" />
                                <asp:Button ID="CancelNewImg" runat="server" Text="取消更新照片" OnClientClick="DoCancelNewImg()" />
                                <div class="descriptiveText" style="display: none">
                                    請選擇上傳您的照片，大小請勿超過<asp:Label ID="AvatarSize" runat="server" /></div>
                            </td>
                        </tr>
                        <tr id="ImgNow" runat="server">
                            <td>
                                <asp:Image ID="ImageNow" runat="server" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>

    <script language="javascript" type="text/javascript">
        $(document).ready(function() {
            //Tab Menu
            $("#tabInfo").addClass("btnDistrict");
            $("#tabList").addClass("btnDistrict now");
            $("#tabVote").addClass("btnDistrict");
            $("#tabFinal").addClass("btnDistrict");

            //Auto Grow
            $("#ctl00_cp_Des").autogrow({
                minHeight: 60,
                lineHeight: 15
            });
        });
        function Hint() {
            $.each($(".descriptiveText"), function(i, n) {
                if ($(this).css("display") == "none") {
                    $(this).show();
                } else {
                    $(this).hide();
                }
            });
        }

        function CheckUpdate() {
            var errMsg = '';

            if ($("#ctl00_cp_NewTopic").val() == '尚未輸入作品名稱' || $("#ctl00_cp_NewTopic").val() == '') {
                errMsg += '請輸入作物名稱!\n';
            }

			if ($("#ctl00_cp_Des").val() == '尚未輸入作品介紹' || $("#ctl00_cp_Des").val() == '') {
                errMsg += '請輸入作品介紹!\n';
            }
			
            if ($("#ctl00_cp_ImgUpload").val() == '') {
                errMsg += '請選擇檔案!\n';
            }

            if ($("#ctl00_cp_NewTopic").val().length > 50) {
                errMsg += '作物名稱長度不能超過50!\n';
            }

            if (errMsg.length > 0) {
                alert(errMsg);
            }
            else {
                $("#Action").val('Update');
                document.forms[0].submit();
            }
        }

        function UpdateAvatar() {
            $("#Action").val('NewImg');
            document.forms[0].submit();
        }

        function Cancel() {
            $("#Action").val('Cancel');
            document.forms[0].submit();
        }

        function DoCancelNewImg() {
            $("#Action").val('CancelNewImg');
        }
    </script>

</asp:Content>

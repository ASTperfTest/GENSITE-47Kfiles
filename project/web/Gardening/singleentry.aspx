<%@ Page Language="C#" ValidateRequest="false" AutoEventWireup="true" CodeFile="singleentry.aspx.cs" Inherits="singleentry" MasterPageFile="Default2.master" %>

<asp:Content ID="Content1" ContentPlaceHolderID="cp" runat="Server">

    <script language="javascript" type="text/javascript" src="js/ui.datepicker.js"></script>

    <script language="javascript" type="text/javascript" src="js/ui.datepicker-zh-TW.js"></script>

    <script type="text/javascript" src="js/jquery.autogrow.js"></script>
    <asp:HiddenField ID="useManagementMasterPage" runat="server" />
    <div id="content">
        <table width="1003px">
            <tr>
                <td class="menu">
                    <input type="hidden" name="PublicCount" id="PublicCount" />
                    <input type="hidden" name="Action" id="Action" value="Non" />
                    <table>
                        <tr>
                            <td id="NewImg" runat="server">
                                <a href="javascript:UpdatePic();">更新照片</a> </td>
                            <td>
                                <a href="javascript:CheckSave();">儲存</a> </td>
                            <td>
                                <a href="javascript:history.go(-1);">取消</a> </td>
                            <td>
                                <a href="javascript:Hint();">欄位說明</a> </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table width="100%" class="grid">
                        <tr>
                            <th>
                                是否公開
                            </th>
                            <td>
                                <asp:RadioButtonList ID="radioIsPublic" runat="server" RepeatDirection="Horizontal">
                                    <asp:ListItem Text="是" Value="true"></asp:ListItem>
                                    <asp:ListItem Text="否" Value="false"></asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                拍攝日期
                            </th>
                            <td>
                                <asp:TextBox ID="DateText" runat="server" contentEditable="false" class="DateText" />
                                <input type="hidden" id="SelectDate" />
                                <div class="descriptiveText2" style="display: none">
                                    請選擇本次上傳照片的拍攝日期</div>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                記錄標題
                            </th>
                            <td>
                                <asp:TextBox ID="TitleText" runat="server" Width="80%" MaxLength="255"></asp:TextBox>
                                <div class="descriptiveText2" style="display: none">
                                    請在此輸入本次紀錄的標題</div>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                種植心得
                            </th>
                            <td align="left">
                                <textarea id="Des" runat="server" cols="80">
                    </textarea><div class="descriptiveText2" style="display: none">
                        請在此紀錄您的種植情況，例如拍照時的土壤，水份，施肥，日照，溫度，蟲害情況等，也可以分享您觀察植物的一些心得</div>
                            </td>
                        </tr>
                        <tr id="ImgUp" runat="server">
                            <th>
                                選取植物照片
                            </th>
                            <td align="left">
                                <asp:FileUpload ID="ImgUpload" runat="server" />
                                <asp:Button ID="CancelNewImg" runat="server" Text="取消更新照片" OnClientClick="DoCancelNewImg()" />
                                <div class="descriptiveText2" style="display: none">
                                    請記得調整您的相機，照片解析度應設為800*600以上，請上傳原始之照片檔案，不要做任何加工或編輯喔！</div>
                            </td>
                        </tr>
                        <tr id="ImgNow" runat="server">
                            <th>
                                植物照片
                            </th>
                            <td align="left">
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

            //TextArea Autogrow
            $("#ctl00_cp_Des").autogrow({
                minHeight: 60,
                lineHeight: 15
            });

            //datepicker
            $("#SelectDate").datepicker({
                showOn: 'both',
                firstDay: 7,
                changeFirstDay: false,
                speed: 'fast',
                buttonImage: 'images/calendar.gif',
                buttonImageOnly: true,
                dateFormat: 'yy/mm/dd',
                onSelect: updateBeginDate
            });

            function updateBeginDate(date) {
                if (date == "") {
                    $('.DateText').val("");
                }
                else {
                    var year = date.substr(0, 4);
                    year = year - 1911;
                    var exceptYear = date.substr(4, 6);
                    $('.DateText').val(year + exceptYear);
                }
            }
        });

        function Hint() {
            $.each($(".descriptiveText2"), function(i, n) {
                if ($(this).css("display") == "none") {
                    $(this).show();
                } else {
                    $(this).hide();
                }
            });
        }

        function CheckSave() {
            var errMsg = '';
            if ($("#PublicCount").val() == "0" &&
            $(":radio:checked")[0].value == "false") {
                errMsg += '最少需要一篇公開日誌!\n';
            }
            if ($("#ctl00_cp_DateText").val() == '') {
                errMsg += '請選擇日期!\n';
            }

            if ($("#ctl00_cp_ImgUpload").val() == '') {
                errMsg += '請選擇檔案!\n';
            }

            if ($("#ctl00_cp_Des").val().length <= 0) {
                errMsg += '請輸入種植心得!\n';
            }

            if (errMsg.length > 0) {
                alert(errMsg);
            }
            else {
                $("#Action").val('Save');
                document.aspnetForm.submit();
            }
        }

        function Cancel() {
            $("#Action").val('Cancel');
            document.aspnetForm.submit();
        }

        function UpdatePic() {
            $("#Action").val('NewImg');
            document.aspnetForm.submit();
        }

        function DoCancelNewImg() {
            $("#Action").val('CancelNewImg');
        }
    </script>

</asp:Content>

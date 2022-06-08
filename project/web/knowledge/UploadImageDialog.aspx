<%@ Page Language="VB" AutoEventWireup="false" CodeFile="UploadImageDialog.aspx.vb" Inherits="UploadImageDialog" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>農業知識入口網 －小知識串成的大力量－</title>
    <base target="_self">
    <script src="../js/jquery-1.3.2.min.js" type="text/javascript"></script>
    <script src="./js/jquery.Jcrop.min.js" type="text/javascript"></script>
    <script src="./js/jquery.Jcrop.js" type="text/javascript"></script>

    <link href="./style/jquery.Jcrop.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" type="text/css" href="/xslgip/style3/css/4seasons.css" />
    
    <script language="javascript" type="text/javascript">
        $(document).ready(function() {
        $('#CropBox').Jcrop({
                onSelect :updateCoords
            });
        });

        function updateCoords(c) {
            $('#SetX').val(c.x);
            $('#SetY').val(c.y);
            $('#SetW').val(c.w);
            $('#SetH').val(c.h);            
        };

        function checkFile(imgFile) {
            var fType = imgFile.value.substring(imgFile.value.lastIndexOf("."), imgFile.value.length);
            if (fType.toLowerCase() != ".jpg" && fType.toLowerCase() != ".jpeg" && fType.toLowerCase() != ".gif")
                alert('只允許上傳JPG或GIF影像檔');
        }
        
        function checkfocuscontent(item) {
            if (item.value == "請輸入圖片說明") {
                item.value = "";
            }
        }

        function isExistFile(control) {
            var fControl = control.id.substring(control.id.lastIndexOf("_"), control.id.length);
            if (fControl == null) { fControl = control.id; }
            if (fControl != null) {
                if (fControl == "btnPreview") {
                    if (document.getElementById('FileUpload').value == "") {
                        alert('請選擇圖片。');
                        return false;
                    }
                }
                else {
                    if (document.getElementById('hidFileName').value == "") {
                        alert('請選擇圖片。');
                        return false;
                    }
                }

                if (fControl == "btnOK") {
                    if (document.getElementById('txtHint').value == "請輸入圖片說明" || document.getElementById('txtHint').value == "") {
                        alert('請輸入圖片說明。');
                        return false;
                    }
                }
                if (fControl == "btnSelect") {
                    if (document.getElementById('SetX').value == ""
                      | document.getElementById('SetY').value == ""
                      | document.getElementById('SetW').value == ""
                      | document.getElementById('SetH').value == ""
                      | document.getElementById('SetW').value == "0"
                      | document.getElementById('SetH').value == "0") {
                            alert('請圈選圖片範圍');
                            return false;
                        }
                }
            }
        }

    </script>
</head>
<body>
    <div class="pantology">
        <div class="head">
        </div>
        <div class="body">
            <h2>圖片裁切上傳</h2>            
            <form id="form1" runat="server">
                <table cellpadding="2">
                    <tr><td colspan="2" width="300px">
                    <asp:FileUpload ID="FileUpload" runat="server" onchange="checkFile(this)" />
                    <asp:Button ID="btnPreview" runat="server" Text="確定上傳" OnClientClick="return isExistFile(this);" /></td>
                    <td rowspan="2"><asp:Label ID="lblText" runat="server" Font-Size="Small" ForeColor="#990000"> 
                              每張圖片檔案大小限制需小於<%=SizeLimit/1024 %>k，檔案格式限定.jpg或.gif。
                              <br/>輸入圖片說明文字可幫助其他人更容易了解。<br/>
			                  圖片上傳後，需先經過網站管理者審核，審核通過者，方會在前台發問或討論頁面呈現。</asp:Label></td></tr>
                    <tr><td colspan="2">
                    圖片寬度：
                    <asp:DropDownList ID="ddlImagePixel" runat="server" AutoPostBack="True">
                        <asp:ListItem Value="0">原始大小</asp:ListItem>
                        <asp:ListItem Value="320">320px</asp:ListItem>
                        <asp:ListItem Value="480">480px</asp:ListItem>
                        <asp:ListItem Value="640">640px</asp:ListItem>
                    </asp:DropDownList><br/>(建議您將圖片寬度調整為640px)                    
                    </td></tr>
                    <tr><td colspan="2">
                    <asp:Button ID="btnSelect" runat="server" Text="確定截取圖片" OnClientClick="return isExistFile(this);" />
                    <asp:Button ID="btnRestore" runat="server" Text="復原原始圖片" OnClientClick="return isExistFile(this);" />
                    </td>
                    <td>
                            <asp:TextBox ID="txtHint" runat="server" onfocus="checkfocuscontent(this);">請輸入圖片說明</asp:TextBox>
                            <asp:Button ID="btnOK" runat="server" Text="完成" OnClientClick="return isExistFile(this);" />
                            <asp:Button ID="btnCancel" runat="server" Text="取消" />                    
                    </td>
                    </tr>
                </table>
                <div>
                    <asp:Image ID="CropBox" runat="server" />
                </div>
                <asp:HiddenField ID="SetX" runat="server" />
                <asp:HiddenField ID="SetY" runat="server" />
                <asp:HiddenField ID="SetW" runat="server" />
                <asp:HiddenField ID="SetH" runat="server" />
                <asp:HiddenField ID="hidFileName" runat="server" />
            </form>
        </div>
        <div class="foot">
        </div>
    </div>
</body>
</html>

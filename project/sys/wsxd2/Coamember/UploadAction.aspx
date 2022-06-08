<%@ Page Language="C#" AutoEventWireup="true" CodeFile="UploadAction.aspx.cs" Inherits="UploadAction" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>個人資料修改 上傳大頭貼</title>
    <style type="text/css">
    .PhotoSize
    {
        max-width:150px;
        max-height:150px;
    }
    </style>
    <script language="javascript" type="text/javascript">
        function GetReturnValue(str) {
            document.getElementById('hidFileName').value = str;
        }  
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table style="table-layout:fixed;">
            <tr style="height:160px;">
                <td valign="center"  style="width:155px;" >
                    <asp:Image ID="MemberPhoto" runat="server" CssClass="PhotoSize" />
                </td>
				<td style="width:280px;" align="Left">
				    <asp:Button ID="uploadImgBtn" CssClass="btn2" runat="server" Text="選擇圖片" OnClientClick=
                        "window.showModalDialog('UploadImageDialog.aspx',self,'resizable:no;scrollbars:Yes;dialogWidth:900px;dialogHeight:600px;center:Yes'); " />
                    <asp:Button ID="deleteBtn" CssClass="btn2" runat="server" Text="刪除圖片" OnClick="deleteBtn_Click" />
                    <!--<asp:FileUpload ID="FileUpload" runat="server" />-->
					<asp:Button ID="UploadBtn" runat="server" Text="存檔" OnClick="UploadBtn_Click" />
					<br />
					 <asp:Label ID="Label1" runat="server" Text="(圖檔限制 1MB )" Font-Size="Small"></asp:Label>               
                    <asp:HiddenField ID="hidFileName" runat="server" />
				</td>
			</tr>
        </table>
    </form>
</body>
</html>

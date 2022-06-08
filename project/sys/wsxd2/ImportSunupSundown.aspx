<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ImportSunupSundown.aspx.cs" Inherits="OfflineExcel" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>匯入日出日落資料</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            &nbsp;<br />
                <asp:FileUpload ID="FileUploadExcel" runat="server" />
                <asp:Button ID="ButtonUpload" runat="server" Text="匯入Excel到資料庫" OnClick="ButtonUpload_Click" />
				<a href='OfflineDown/TemplateSunData.xls'>下載範本</a>
            </div>
    </form>
</body>
</html>

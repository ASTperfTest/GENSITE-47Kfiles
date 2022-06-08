<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ExportPage.aspx.cs" Inherits="ExportPDFFile_ExportPage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<asp:Button ID="btnCalTotal" runat="server" Text="計算檔案筆數" 
        onclick="btnCalTotal_Click" />
&nbsp;
      <asp:Label ID="labelTotalFileCount" runat="server"></asp:Label>
      <br />
      <br />
<asp:Button ID="btnCopy" runat="server" Text="取出pdf檔案" onclick="btnCopy_Click" />
<asp:Label ID="labelExportResult" runat="server"></asp:Label>
      <br />
      <br />
      <asp:Button ID="Button1" runat="server" onclick="Button1_Click" Text="取得路徑" />
      <asp:Label ID="labelFilePath" runat="server"></asp:Label>
    </div>
    </form>
</body>
</html>

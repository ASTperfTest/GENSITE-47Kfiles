<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DataImport.aspx.cs" Inherits="maToolKits_DataImport" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>農業百年大事紀 - 資料管理</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <link type="text/css" rel="stylesheet" href="../css/form.css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <div id="FuncName">
            <h1>
                農業百年大事紀 - 資料管理
            </h1>
            <div id="Nav">
            </div>
            <div id="ClearFloat">
            </div>
        </div>
        <div id="Event">
            <div id="DataMaintenace">
                <asp:GridView ID="GridView1" runat="server">
                </asp:GridView>
                <asp:FileUpload ID="fileCSV" runat="server" />
                <asp:Button runat="server" ID="btnImport" CssClass="cbutton" Text="匯入資料" 
                    onclick="btnImport_Click" />

                <asp:Button runat="server" ID="btnClearDB" CssClass="cbutton" Text="刪除全資料" 
                    OnClientClick="return confirm('確定將資料庫資料刪除?');" onclick="btnClearDB_Click" />
            </div>
        </div>
    </div>
    </form>
</body>
</html>

<%@ Page Language="C#" AutoEventWireup="true" CodeFile="bgMusicEdit.aspx.cs" Inherits="bgMusic_bgMusicEdit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link rel="stylesheet" href="/inc/setstyle.css" />
    <link href="../css/list.css" rel="stylesheet" type="text/css" />
    <link href="../css/layout.css" rel="stylesheet" type="text/css" />
    <title>背景音樂管理</title>
    <script type="text/javascript" language="javascript">
        
        function RetrunMessage(msg) {
            alert(msg);
            window.location = "bgMusicList.aspx";
        }

    </script>

</head>
<body>
    <form id="form1" runat="server">
    <div id="FuncName">
        <h1><asp:Label ID="Lbltitle" runat="server"></asp:Label></h1>
        <div id="Nav">
        <a href="bgMusicList.aspx">回清單</a>
        </div>
        <div id="ClearFloat">
        </div>
    </div>
    <div style=" margin-left:10%">
        <table border="0" id="ListTable" style="width:60%" cellspacing="1" cellpadding="0" align="center"  >
            <tr>
                <td style=" background-color:#D0F1BD; width:100px; text-align:right">曲目標題
                </td>
                <td>
                    <asp:TextBox ID="txttitle" runat="server"></asp:TextBox>(限10個字元)
                </td>
            </tr>
            <tr>
                <td style="background-color: #D0F1BD; width: 100px; text-align: right">
                    上傳檔案
                </td>
                <td>
                    <asp:FileUpload ID="FUMusic" runat="server" />&nbsp;
                    <asp:Label ID="lbfilename" runat="server"></asp:Label>
                </td>
            </tr>
        </table>
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr>
                <td align="center">
                    <asp:Button ID="Btsave" runat="server" Text="確定新增" onclick="Btsave_Click" />
                    <asp:Button ID="Btcancel" runat="server" Text="取消" onclick="Btcancel_Click" />
                </td>
            </tr>
        </table>
    </div>
    
    </form>
</body>
</html>

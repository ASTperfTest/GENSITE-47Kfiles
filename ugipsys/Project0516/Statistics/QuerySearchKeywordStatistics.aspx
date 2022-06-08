<%@ Page Language="C#" AutoEventWireup="true" CodeFile="QuerySearchKeywordStatistics.aspx.cs"
    Inherits="QuerySearchKeywordStatistics" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>搜索關鍵字</title>

    <script type="text/javascript" src="../js/jquery-1.3.2.min.js"></script>

    <script type="text/javascript" src="../js/datepicker.js"></script>

    <link rel="Stylesheet" href="../css/datepicker.css" />
    <link type="text/css" href="../css/jquery.css" rel="stylesheet" />

    <script type="text/javascript" language="JavaScript">


        $(document).ready(function(e) {
            $("#tboxDateStart").datepicker({
                dateFormat: "yy/mm/dd"
            });
        })
        $(document).ready(function(e) {
            $("#tboxDateEnd").datepicker({
                dateFormat: "yy/mm/dd"
            });
        })
    
    </script>

</head>
<body>
    <form id="form1" runat="server">
    查詢關鍵字：<asp:TextBox ID="tboxTag" runat="server"></asp:TextBox><br />
    開始日期：<asp:TextBox ID="tboxDateStart" runat="server"></asp:TextBox>～ 結束日期：<asp:TextBox
        ID="tboxDateEnd" runat="server"></asp:TextBox><br />
    查詢數量：<asp:TextBox ID="tboxTagNumStart" runat="server"></asp:TextBox>～
    <asp:TextBox ID="tboxTagNumEnd" runat="server"></asp:TextBox>
    <br />
    排序依：<asp:DropDownList ID="sortOrder" runat="server">
    <asp:ListItem Text="查詢數量" Selected="True" Value="USED_COUNT"></asp:ListItem>
    <asp:ListItem Text="最後查詢日期"  Value="LAST_USED"></asp:ListItem>
    <asp:ListItem Text="關鍵字"  Value="DISPLAY_NAME"></asp:ListItem>
    </asp:DropDownList>
    <asp:DropDownList ID="orderBy" runat="server">
    <asp:ListItem Text="遞減" Selected="True" Value="DESC"></asp:ListItem>
    <asp:ListItem Text="遞增"  Value="ASC"></asp:ListItem>
    </asp:DropDownList>
    <br />
    <asp:Button ID="btnQuery" runat="server" Text="查詢" OnClick="btnQuery_Click" />
    <br />
    <br />
    <hr />
    <br />
    <asp:Label ID="StatisticsTitle" runat="server" Text="統計結果" Font-Bold="True" Visible="False"></asp:Label><br />
    <asp:GridView ID="GridViewOfQueryResult" runat="server" Height="30px" Width="800px"
        CellPadding="4" ForeColor="#333333" GridLines="Both" BorderStyle="Dashed">
        <RowStyle BackColor="#EFF3FB" HorizontalAlign="Center" VerticalAlign="Middle" />
        <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
        <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
        <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
        <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
        <EditRowStyle BackColor="#2461BF" />
        <AlternatingRowStyle BackColor="White" />
    </asp:GridView>
    <asp:Label ID="lblNodata" runat="server" Text="Label" Visible="false"></asp:Label>
    </form>
</body>
</html>

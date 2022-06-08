<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SubjectClickStatistics.aspx.vb"
    Inherits="Statistics_SubjectClickStatistics" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>各主題館瀏覽次數統計</title>

    <script type="text/javascript" src="../js/jquery-1.3.2.min.js"></script>

    <script type="text/javascript" src="../js/datepicker.js"></script>

    <link rel="Stylesheet" href="../css/datepicker.css" />
    <link type="text/css" href="../css/jquery.css" rel="stylesheet" />

    <script type="text/javascript" language="JavaScript">
 
    
    $(document).ready(function(e){
		$("#txtStartDate").datepicker({
        dateFormat:"yy/mm/dd"
        }); 
    })
    $(document).ready(function(e){
        $("#txtEndDate").datepicker({
        dateFormat:"yy/mm/dd"
        });             
    })
    
    </script>

</head>
<body>
    <form id="form1" runat="server">
        各主題館瀏覽次數統計<br />
        <br />
        <div>
            開始日期：<asp:TextBox ID="txtStartDate" runat="server"></asp:TextBox>
            結束日期：<asp:TextBox ID="txtEndDate" runat="server"></asp:TextBox>
            <asp:Button ID="btnQuery" runat="server" Text="查詢" />
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
            <br />
        </div>
    </form>
</body>
</html>

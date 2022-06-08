<%@ Page Language="VB" AutoEventWireup="false" CodeFile="MpClickStatistics.aspx.vb" Inherits="Statistics_MpClickStatistics" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>各單元瀏覽次數統計</title>
    
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
首頁總瀏覽數相加明細

    <form id="form1" runat="server">
        <div>
            
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
            本頁所使用的首頁瀏覽數並非是統計->各單元瀏覽內首頁瀏覽資料<br/>
            而是最初設計的首頁計數
        </div>
    </form>
</body>
</html>
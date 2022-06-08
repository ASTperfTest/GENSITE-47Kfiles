<%@ Page Language="C#" AutoEventWireup="true" CodeFile="QueryPeriodUserAgeStatistics.aspx.cs"
    Inherits="QueryPeriodUserAgeStatistics" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>登入年齡分佈統計</title>

    <script type="text/javascript" src="js/jquery-1.3.2.min.js"></script>

    <script type="text/javascript" src="js/datepicker.js"></script>

    <link rel="Stylesheet" href="css/datepicker.css" />
    <link type="text/css" href="css/jquery.css" rel="stylesheet" />

    <script type="text/javascript" language="JavaScript">
 
    
    $(document).ready(function(e){
		$("#TextBoxBeginDate").datepicker({
        dateFormat:"yy/mm/dd"
        }); 
    })
    $(document).ready(function(e){
        $("#TextBoxEndDate").datepicker({
        dateFormat:"yy/mm/dd"
        });             
    })  

    </script>

</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Label ID="Label2" runat="server" Text="開始日期：" ></asp:Label>
            <asp:TextBox ID="TextBoxBeginDate" runat="server" MaxLength="10" ToolTip="滑鼠點一下，可使用小月曆選擇日期" ValidationGroup="query" Width="100px"></asp:TextBox>&nbsp;
           
            <asp:Label ID="Label3" runat="server" Text="結束日期：" ></asp:Label>
            <asp:TextBox ID="TextBoxEndDate" runat="server" MaxLength="10"  ToolTip="滑鼠點一下，可使用小月曆選擇日期" ValidationGroup="query" Width="100px"></asp:TextBox>&nbsp;
            
            <asp:Button ID="Query" runat="server" Text="開始查詢" OnClick="Query_Click" ValidationGroup="query"  />
            <br />
<br />
<hr />
            <br />
            <asp:Label ID="StatisticsTitle" runat="server" Text="統計結果" Font-Bold="True" Visible="False"></asp:Label><br />
            <asp:GridView ID="GridViewOfQueryResult" runat="server" Height="30px" Width="400px" CellPadding="4" ForeColor="#333333" GridLines="Both" BorderStyle="Dashed">
                <RowStyle BackColor="#EFF3FB" HorizontalAlign="Center" VerticalAlign="Middle" />
                <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#2461BF" ForeColor="White" HorizontalAlign="Center" />
                <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="True" ForeColor="#333333" />
                <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                <EditRowStyle BackColor="#2461BF" />
                <AlternatingRowStyle BackColor="White" />
            </asp:GridView>
            <br />
        </div>
    </form>
    <div id="testdiv1" style="position: absolute; visibility: hidden; background-color: white;">
    </div>
</body>
</html>

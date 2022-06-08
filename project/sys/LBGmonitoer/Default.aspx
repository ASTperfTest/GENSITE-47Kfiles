<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="監看系統._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <asp:Label ID="Label1" runat="server" Text="姓名"></asp:Label>
        <asp:TextBox ID="txtName" runat="server" Width="217px"></asp:TextBox>
        <br />
        <asp:Label ID="Label2" runat="server" Text="暱稱"></asp:Label>
        <asp:TextBox ID="txtNickName" runat="server" Width="218px"></asp:TextBox>
        <br />
        <asp:Label ID="Label3" runat="server" Text="分數區間"></asp:Label>
        <asp:TextBox ID="txtScoreLowBound" runat="server" Width="80px"></asp:TextBox>
        <asp:Label ID="Label4" runat="server" Text="～"></asp:Label>
        <asp:TextBox ID="txtScoreHighBound" runat="server" Width="92px"></asp:TextBox>
        <asp:Button ID="Button1" runat="server" onclick="btn_Refresh_Click" Text="查詢" />
        <asp:Button ID="btn_Reset" runat="server" onclick="btn_Reset_Click" 
            Text="重設查詢條件" />
        <br />
        <asp:Label ID="Label5" runat="server" Text="Label"></asp:Label>
    
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" 
            EmptyDataText="RR" onpageindexchanged="GridView1_PageIndexChanging" 
            onpageindexchanging="GridView1_PageIndexChanging1" 
            onsorting="GridView1_Sorting" ShowFooter="True" Width="100%" 
            AutoGenerateColumns="False" DataSourceID="SqlDataSource">
            <Columns>
                <asp:BoundField DataField="REALNAME" HeaderText="真實姓名" 
                    SortExpression="REALNAME" />
                <asp:BoundField DataField="NICKNAME" HeaderText="暱稱" 
                    SortExpression="NICKNAME" />
                <asp:BoundField DataField="TEL" HeaderText="電話" SortExpression="TEL" />
                <asp:BoundField DataField="EMAIL" HeaderText="e-mail" SortExpression="EMAIL" />
                <asp:BoundField DataField="ADDRESS" HeaderText="地址" SortExpression="ADDRESS" />
                <asp:BoundField DataField="TOTALSTAR" HeaderText="總得星" 
                    SortExpression="TOTALSTAR" />
                <asp:BoundField DataField="TOTALSCORE" HeaderText="總得分" 
                    SortExpression="TOTALSCORE" />
                <asp:BoundField DataField="PLANT_A_SCORE" HeaderText="菜葉甘藷分數" 
                    SortExpression="PLANT_A_SCORE">
                    <ItemStyle BackColor="#FF9933" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_A_STAR" HeaderText="菜葉甘藷星數" 
                    SortExpression="PLANT_A_STAR">
                    <ItemStyle BackColor="#FF9933" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_A_CURRENT_SCORE" HeaderText="菜葉甘藷目前得分" 
                    SortExpression="PLANT_A_CURRENT_SCORE">
                    <ItemStyle BackColor="#FF9933" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_A_PSEUDO_STAR" HeaderText="菜葉甘藷預期得星" 
                    SortExpression="PLANT_A_PSEUDO_STAR">
                    <ItemStyle BackColor="#FF9933" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_B_SCORE" HeaderText="孤挺花分數" 
                    SortExpression="PLANT_B_SCORE">
                    <ItemStyle BackColor="#66FFCC" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_B_STAR" HeaderText="孤挺花星數" 
                    SortExpression="PLANT_B_STAR">
                    <ItemStyle BackColor="#66FFCC" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_B_CURRENT_SCORE" HeaderText="孤挺花目前得分" 
                    SortExpression="PLANT_B_CURRENT_SCORE">
                    <ItemStyle BackColor="#66FFCC" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_B_PSEUDO_STAR" HeaderText="孤挺花預期得星" 
                    SortExpression="PLANT_B_PSEUDO_STAR">
                    <ItemStyle BackColor="#66FFCC" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_C_SCORE" HeaderText="毛豆分數" 
                    SortExpression="PLANT_C_SCORE">
                    <ItemStyle BackColor="Lime" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_C_STAR" HeaderText="毛豆星數" 
                    SortExpression="PLANT_C_STAR">
                    <ItemStyle BackColor="Lime" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_C_CURRENT_SCORE" HeaderText="毛豆目前得分" 
                    SortExpression="PLANT_C_CURRENT_SCORE">
                    <ItemStyle BackColor="Lime" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_C_PSEUDO_STAR" HeaderText="毛豆預期得星" 
                    SortExpression="PLANT_C_PSEUDO_STAR">
                    <ItemStyle BackColor="Lime" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_D_SCORE" HeaderText="海芋分數" 
                    SortExpression="PLANT_D_SCORE">
                    <ItemStyle BackColor="Yellow" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_D_STAR" HeaderText="海芋星數" 
                    SortExpression="PLANT_D_STAR">
                    <ItemStyle BackColor="Yellow" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_D_CURRENT_SCORE" HeaderText="海芋目前得分" 
                    SortExpression="PLANT_D_CURRENT_SCORE">
                    <ItemStyle BackColor="Yellow" />
                </asp:BoundField>
                <asp:BoundField DataField="PLANT_D_PSEUDO_STAR" HeaderText="海芋預期得星" 
                    SortExpression="PLANT_D_PSEUDO_STAR">
                    <ItemStyle BackColor="Yellow" />
                </asp:BoundField>
            </Columns>
            <HeaderStyle BackColor="#33CCFF" />
            <AlternatingRowStyle BackColor="#99FF33" />
        </asp:GridView>
    
        <asp:SqlDataSource ID="SqlDataSource" runat="server" 
            ConnectionString="<%$ ConnectionStrings:ConnectionString2 %>" 
            onselecting="SqlDataSource_Selecting" 
            
            SelectCommand="SELECT [UID], [REALNAME], [NICKNAME], [TEL], [EMAIL], [ADDRESS], [TOTALSTAR], [TOTALSCORE], [PLANT_A_SCORE], [PLANT_A_STAR], [PLANT_A_PSEUDO_STAR], [PLANT_A_CURRENT_SCORE], [PLANT_B_SCORE], [PLANT_B_STAR], [PLANT_B_PSEUDO_STAR], [PLANT_D_SCORE], [PLANT_C_CURRENT_SCORE], [PLANT_D_STAR], [PLANT_D_PSEUDO_STAR], [PLANT_D_CURRENT_SCORE], [PLANT_B_CURRENT_SCORE], [PLANT_C_SCORE], [PLANT_C_STAR], [PLANT_C_PSEUDO_STAR] FROM [LBG_SCORE_SUMMARY] ORDER BY [TOTALSTAR] DESC, [TOTALSCORE] DESC" 
            ProviderName="<%$ ConnectionStrings:ConnectionString2.ProviderName %>">
        </asp:SqlDataSource>
    
    </div>
    <asp:Button ID="btn_Refresh" runat="server" onclick="btn_Renew_Refresh_Click" 
        Text="更新" />
    <asp:Button ID="btn_Excel" runat="server" onclick="btn_Excel_Click" 
        Text="匯出Excel檔" />
    </form>
</body>
</html>

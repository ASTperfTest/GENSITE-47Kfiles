<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="kpi_set_index.aspx.vb" Inherits="kpi_set_index" title="index" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
   <table border="0" width="100%" cellspacing="0" cellpadding="0">
	  
	  <tr>
	    <td width="100%" colspan="2" style="height: 19px">
	      
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  

</table> 
   
   <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnString %>"
        SelectCommand="SELECT [Rank0], [Rank1], [Rank2] FROM [kpi_set_ind]">
    </asp:SqlDataSource>

    <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="create.aspx" BackColor="White" BorderColor="White" ForeColor="CornflowerBlue">新增</asp:HyperLink>
    <asp:GridView CssClass="" ID="GridView1" runat="server" AutoGenerateColumns="False" DataSourceID="SqlDataSource1" CellPadding="4" ForeColor="#333333" GridLines="None" Height="171px" Width="1230px">
        <Columns>
            <asp:HyperLinkField DataTextField="Rank1" HeaderText="排行榜" DataNavigateUrlFormatString="kpi_ed1.aspx?id={0}" DataNavigateUrlFields="Rank0" NavigateUrl="kpi_ed1.aspx" />
            <asp:HyperLinkField HeaderText="設定" Text="詳細設定" DataNavigateUrlFields="Rank0" DataNavigateUrlFormatString="kpi_set_scored.aspx?id={0}" />
        
        
        </Columns>
        <FooterStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
        <RowStyle BackColor="#FFFBD6" ForeColor="#333333" />
        <PagerStyle BackColor="#FFCC66" ForeColor="#333333" HorizontalAlign="Center" />
        <SelectedRowStyle BackColor="#FFCC66" Font-Bold="True" ForeColor="Navy" />
        <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="White" />
        <AlternatingRowStyle BackColor="White" />
    </asp:GridView>
   

    <br /><hr /> 


    
   
 
    


</asp:Content>



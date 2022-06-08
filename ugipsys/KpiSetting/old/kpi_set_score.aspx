<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="kpi_set_score.aspx.vb" Inherits="kpi_set_score" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    
    
 
     
    
    
    &nbsp;<asp:Label ID="Label6" runat="server" ForeColor="Black" Text="Label" Width="267px"></asp:Label>
    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
    <asp:GridView ID="GridView1" runat="server" CellPadding="4" GridLines="Horizontal" AutoGenerateColumns="False" BackColor="White" BorderColor="#336666" BorderStyle="Double" BorderWidth="3px">
        <FooterStyle BackColor="White" ForeColor="#333333" />
        <RowStyle BackColor="White" ForeColor="#333333" />
        <Columns>
            <asp:BoundField DataField="名稱" HeaderText="名稱" />
            <asp:BoundField DataField="分數" HeaderText="分數" />
            <asp:HyperLinkField DataNavigateUrlFields="Rank0_2,Rank0_4" DataNavigateUrlFormatString="score_ed.aspx?id={0}&amp;sid={1}"
                Text="設定" />
        </Columns>
        <PagerStyle BackColor="#336666" ForeColor="White" HorizontalAlign="Center" />
        <SelectedRowStyle BackColor="#339966" Font-Bold="True" ForeColor="White" />
        <HeaderStyle BackColor="#336666" Font-Bold="True" ForeColor="White" />
    </asp:GridView>
    <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="create1.aspx" Visible="False">新增</asp:HyperLink>
    
    <hr>
  
	<TABLE width="100%" id="ListTable">
  <tr>
	    <td width="50%" align="left" nowrap class="FormName">會員KPI記錄分數設定&nbsp;
		</td>
		<td width="50%" class="FormLink" align="right">
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A>			
		</td>	
  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
<CENTER>
                        
                        
    <TR>
      <Th width="20%">
                    <asp:Label ID="Label0" runat="server" Text="Label" Width="246px"></asp:Label>：</Th>
      <td style="width: 247px">
          &nbsp;<asp:TextBox id="TextBox1" runat="server" Width="85px"></asp:TextBox>%</td>：
      <TD class="eTableContent">
        </TD>
    </TR>
    <TR>
  <Th style="height: 26px">
      <asp:Label ID="Label1" runat="server" Text="Label" Width="246px"></asp:Label>：</Th>
  <TD class="eTableContent" style="height: 26px">
      &nbsp;<asp:TextBox ID="TextBox2" runat="server" Width="85px"></asp:TextBox>%</TD>
</TR>
    <TR>
      <Th>
          <asp:Label ID="Label2" runat="server" Text="Label" Width="246px"></asp:Label>：</Th>
      <TD class="eTableContent">
          &nbsp;<asp:TextBox ID="TextBox3" runat="server" Width="82px"></asp:TextBox>%</TD>
    </TR>
    <TR>
      <Th>
          <asp:Label ID="Label3" runat="server" Text="Label" Width="246px"></asp:Label>：</Th>
      <TD class="eTableContent">
          &nbsp;<asp:TextBox ID="TextBox4" runat="server" Width="82px"></asp:TextBox>%</TD>
    </TR>
    <TR>
      <Th>
          <asp:Label ID="Label4" runat="server" Text="Label" Width="246px"></asp:Label>：</Th>
      <TD class="eTableContent">
          &nbsp;<asp:TextBox ID="TextBox5" runat="server" Width="82px"></asp:TextBox>%</TD>
    </TR>

</CENTER>

<tr>
 <td align="center">     
        <input name="button4" type="button" class="cbutton" onClick="location.href='kpiListset.htm'" value="編修存檔">
        <input type=button value ="重　設" class="cbutton" onClick="resetForm()">
        <input type=button value ="取　消" class="cbutton" onClick="resetForm()"></td>
</tr>
</table>













</asp:Content>


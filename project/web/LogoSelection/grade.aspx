<%@ Page Language="C#" AutoEventWireup="true" CodeFile="grade.aspx.cs" Inherits="grade" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	 <title><%=title%></title>
     <script type="text/javascript">
	     function opennew(url) {
             window.open(url, '', 'width=700,height=400,resizable=yes,toolbar=no');
         }
    </script>
	<link rel="stylesheet" type="text/css" href="./css/css.css">
</head>
<body>
    <form id="form1" runat="server">
    <div id ="gradetable" runat ="server">
        <asp:DataList ID="DataList1" runat="server" OnEditCommand="DataList1_EditCommand" OnCancelCommand="DataList1_CancelCommand" OnUpdateCommand="DataList1_UpdateCommand"  CellPadding="4" ForeColor="#333333">
        <HeaderTemplate >
                <table id="list" cellpadding=0 cellspacing=0>
                <tr>
                    <th>參賽人員 </th>
                    <th>作品 </th>
                    <th>文件備齊 </th>
                    <th>名次 </th>
                    <th>&nbsp;</th>
                </tr>
               
            </HeaderTemplate>
            <ItemTemplate >
                <tr>
                <td> <%#Eval ("CREATOR_DISPLAY_NAME") %> </td>
                <td><a href="javascript:opennew('<%# "/logoselection/fullimage.aspx?fileName="+Eval ("CREATOR")+".jpg"%>')"><img src =' <%# "/logoselection/download.aspx?fileName="+Eval ("CREATOR")+".jpg"%>' Width="200" Height="100" border="0"></a></td>
				<td align="center"> <asp:CheckBox ID="CheckBoxShow" runat="server" Checked=' <%#Eval ("COMPLETED") %>' Enabled="false"/> </td>
                <td> <%#Eval ("SORT_ORDER") %>&nbsp; </td>
                <td> <asp:LinkButton ID="LinkButton1" runat="server" CommandName ="edit">編輯 </asp:LinkButton> </td>
                </tr>
            </ItemTemplate> 
            <EditItemTemplate >
                <tr>
                    <td> <%#Eval ("CREATOR") %> </td>
                    <td><a href="javascript:opennew('<%# "/logoselection/fullimage.aspx?fileName="+Eval ("CREATOR")+".jpg"%>')"><img src =' <%# "/logoselection/download.aspx?fileName="+Eval ("CREATOR")+".jpg"%>' Width="200" Height="100" border="0"></a></td>
                    <td> <asp:CheckBox ID="CheckBoxEdit" runat="server" Checked=' <%#Eval ("COMPLETED") %>'/> </td>
                    <td><asp:TextBox ID="TextRank" runat="server" Text='<%#Eval ("SORT_ORDER") %>' Columns='2'> </asp:TextBox> </td>
                    <td> <asp:LinkButton id="lblUpdata1" runat="server" CommandName="update">更新 </asp:LinkButton>
                    <asp:LinkButton id="lblCancel1" runat="server" CommandName="cancel">取消 </asp:LinkButton> </td>   
                </tr>
            </EditItemTemplate> 
            <FooterTemplate>
                </table>
            </FooterTemplate>
            </asp:DataList>
    </div>
    </form>
</body>
</html>

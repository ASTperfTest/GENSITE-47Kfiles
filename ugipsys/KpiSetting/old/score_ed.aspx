<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="score_ed.aspx.vb" Inherits="score_ed" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    &nbsp;
    <table border="1" cellpadding="5">
            <tr>
                <th colspan="2">
                    �ק�
                </th>
            </tr>
            <tr>
                <td style="text-align: center">
                    �W��
                </td>
                <td style="width: 247px">
                    <asp:Label ID="Label1" runat="server" Text="Label" Width="246px"></asp:Label></td>
            </tr>
            <tr>
                <td style="text-align: center">
                   ����
                </td>
                <td style="width: 247px">
                    <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="TextBox2"
                        Display="Dynamic" SetFocusOnError="True">���i�H�ť�!</asp:RequiredFieldValidator>
                    
                   
                </td>
            </tr>
            
    
        </table>
    &nbsp;
    <asp:Button ID="Button1" runat="server" Text="�ק�" OnClientClick="kpi_set_score.aspx"  />
    <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/kpi_set_index.aspx">�^����</asp:HyperLink>
   <script type="text/javascript"> 

    </script>

    <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ControlToValidate="TextBox2"
        ErrorMessage="�ХX�J�Ʀr" ValidationExpression="\d*"></asp:RegularExpressionValidator>
</asp:Content>


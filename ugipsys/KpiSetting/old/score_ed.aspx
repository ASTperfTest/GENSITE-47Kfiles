<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="score_ed.aspx.vb" Inherits="score_ed" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    &nbsp;
    <table border="1" cellpadding="5">
            <tr>
                <th colspan="2">
                    修改
                </th>
            </tr>
            <tr>
                <td style="text-align: center">
                    名稱
                </td>
                <td style="width: 247px">
                    <asp:Label ID="Label1" runat="server" Text="Label" Width="246px"></asp:Label></td>
            </tr>
            <tr>
                <td style="text-align: center">
                   分數
                </td>
                <td style="width: 247px">
                    <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="TextBox2"
                        Display="Dynamic" SetFocusOnError="True">不可以空白!</asp:RequiredFieldValidator>
                    
                   
                </td>
            </tr>
            
    
        </table>
    &nbsp;
    <asp:Button ID="Button1" runat="server" Text="修改" OnClientClick="kpi_set_score.aspx"  />
    <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="~/kpi_set_index.aspx">回首頁</asp:HyperLink>
   <script type="text/javascript"> 

    </script>

    <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ControlToValidate="TextBox2"
        ErrorMessage="請出入數字" ValidationExpression="\d*"></asp:RegularExpressionValidator>
</asp:Content>


<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="create1.aspx.vb" Inherits="create1" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
<table border="1" cellpadding="5">
            <tr>
                <th colspan="2">
                    新增
                </th>
            </tr>
            <tr>
                <td style="text-align: center" class="FormName">
                   ctNode值
                </td>
                <td>
                    <asp:TextBox ID="TextBox1" runat="server" ></asp:TextBox>
                <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1"
                        Display="Dynamic" SetFocusOnError="True">不可以空白!</asp:RequiredFieldValidator>
                
                </td>
            </tr>
            <tr>
                <td style="text-align: center" class="FormName">
                    分數
                   
                </td>
                <td>
                    <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
                <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="TextBox2"
                        Display="Dynamic" SetFocusOnError="True">不可以空白!</asp:RequiredFieldValidator>
                
                
                </td>
            </tr>
            
        </table>
    &nbsp;&nbsp;
    <asp:Button ID="Button1" runat="server" Text="新增" />
    
    











</asp:Content>


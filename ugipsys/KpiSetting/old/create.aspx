<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="create.aspx.vb" Inherits="create" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">


<table border="1" cellpadding="5">
            <tr>
                <th colspan="2">
                    �s�W
                </th>
            </tr>
            <tr>
                <td style="text-align: center" class="FormName">
                    �Ʀ�]�W��id
                </td>
                <td>
                    <asp:TextBox ID="TextBox1" runat="server" ></asp:TextBox>
                <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1"
                        Display="Dynamic" SetFocusOnError="True">���i�H�ť�!</asp:RequiredFieldValidator>
                
                </td>
            </tr>
            <tr>
                <td style="text-align: center" class="FormName">
                    �Ʀ�]�W��
                   
                </td>
                <td>
                    <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
                <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="TextBox2"
                        Display="Dynamic" SetFocusOnError="True">���i�H�ť�!</asp:RequiredFieldValidator>
                
                
                </td>
            </tr>
            <tr>
                <td style="text-align: center" class="FormName">
                    ���A�}��
                </td>
                <td style ="text-align:center  " class="FormName" >
                    <select id="Select1"  runat="server">
                               <option selected="selected">����</option>
                               <option>�}��</option>
                    </select>
                </td>
            </tr>
    
        </table>
    &nbsp;&nbsp;
    <asp:Button ID="Button1" runat="server" Text="�s�W" />
    <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="kpi_set_index.aspx">�^����</asp:HyperLink>


</asp:Content>


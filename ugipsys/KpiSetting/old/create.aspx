<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="create.aspx.vb" Inherits="create" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">


<table border="1" cellpadding="5">
            <tr>
                <th colspan="2">
                    新增
                </th>
            </tr>
            <tr>
                <td style="text-align: center" class="FormName">
                    排行榜名稱id
                </td>
                <td>
                    <asp:TextBox ID="TextBox1" runat="server" ></asp:TextBox>
                <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="TextBox1"
                        Display="Dynamic" SetFocusOnError="True">不可以空白!</asp:RequiredFieldValidator>
                
                </td>
            </tr>
            <tr>
                <td style="text-align: center" class="FormName">
                    排行榜名稱
                   
                </td>
                <td>
                    <asp:TextBox ID="TextBox2" runat="server"></asp:TextBox>
                <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ControlToValidate="TextBox2"
                        Display="Dynamic" SetFocusOnError="True">不可以空白!</asp:RequiredFieldValidator>
                
                
                </td>
            </tr>
            <tr>
                <td style="text-align: center" class="FormName">
                    狀態開啟
                </td>
                <td style ="text-align:center  " class="FormName" >
                    <select id="Select1"  runat="server">
                               <option selected="selected">關閉</option>
                               <option>開啟</option>
                    </select>
                </td>
            </tr>
    
        </table>
    &nbsp;&nbsp;
    <asp:Button ID="Button1" runat="server" Text="新增" />
    <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="kpi_set_index.aspx">回首頁</asp:HyperLink>


</asp:Content>


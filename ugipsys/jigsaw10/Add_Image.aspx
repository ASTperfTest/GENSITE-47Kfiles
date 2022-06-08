<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Add_Image.aspx.cs" Inherits="Add_Image" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>功能管理 / 農漁生產地圖 / 農漁作物管理</title>
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <link type="text/css" rel="stylesheet" href="../css/setstyle.css" />
    <script language="javascript" type="text/javascript">
        // 實作全選
        function Select_All(spanChk) {
            elm = document.forms[0];
            for (i = 0; i <= elm.length - 1; i++) {
                if (elm[i].type == "checkbox" && elm[i].id != spanChk.id) {
                    if (!elm.elements[i].checked)
                        elm.elements[i].click();
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr>
                <td colspan="3" width="100%" class="FormName" align="left">
                    農漁作物管理&nbsp;<font size="2">【上傳編輯作物照片】</font>
                </td>
            </tr>
            <tr>
                <td width="100" colspan="2">
                    <hr />
                    <font color="red">*</font><font size="2">代表必輸入項目</font>
                </td>
            </tr>
            <tr>
                <td width="80%" height="230" align="center" valign="top" colspan="2">
                    <table id="ListTable">
                        <tbody>
                            <tr>
                                <th>
                                    <font color="red">*</font>圖檔名稱：
                                </th>
                                <td class="eTableContent">
                                    <asp:TextBox ID="txtName" runat="server"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <th>
                                    <font color="red">*</font>圖檔上傳：
                                </th>
                                <td>
                                    <asp:FileUpload ID="fileupload1" runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td width="50%">
                                </td>
                                <td width="45%">
                                </td>
                                <td width="5%" align="right">
                                    <asp:Button ID="btnSubmit" runat="server" class="cbutton" Text="上傳" OnClick="btnSubmit_Click" />
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <table>
                        <tbody>
                            <tr>
                                <td>
                                    <asp:GridView ID="gvView" runat="server" EnableModelValidation="True"  AutoGenerateColumns="false">
                                    <Columns>
                                            <asp:TemplateField>
                                                <HeaderTemplate>
                                                    <input type="button" name="btnSelectAll" value="全選" class="cbutton" onclick="Select_All(this);" />
                                                    <asp:Button ID="btnDelete" runat="server" Text="刪除" CssClass="cbutton" OnClick="btnDelete_Click"
                                                        OnClientClick="return confirm('確認刪除?');" />
                                                </HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:CheckBox ID="CheckBox1" runat="server" />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Center" BackColor="#d0f1bd" />
                                                <ItemStyle HorizontalAlign="Center" BackColor="#d0f1bd" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="作物圖片">
                                                <ItemTemplate>
                                                    <a href="../public/Data/jigsaw/<%=Request["item"] %>/<%# Eval("NFileName") %>" target="_blank">
                                                        <img id="img1" alt="<%# Eval("aTitle") %>" src="../public/Data/jigsaw/<%=Request["item"] %>/<%# Eval("NFileName") %>"
                                                            width="200px" /></a>
                                                </ItemTemplate>
                                                <HeaderStyle BackColor="#d0f1bd" />
                                            </asp:TemplateField>
                                            <asp:TemplateField HeaderText="圖說">
                                                <ItemTemplate>
                                                    <asp:Label ID="labContent" runat="server" Text='<%# Eval("aTitle") %>'></asp:Label>
                                                </ItemTemplate>
                                                <HeaderStyle BackColor="#d0f1bd" />
                                            </asp:TemplateField>
                                        </Columns>
                                    </asp:GridView>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <table>
                        <tbody>
                            <tr>
                                <td align="center">
                                    <asp:Button ID="btnClose" runat="server" Text="關閉" CssClass="cbutton" OnClientClick="javascript:void(window.close());" />
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>

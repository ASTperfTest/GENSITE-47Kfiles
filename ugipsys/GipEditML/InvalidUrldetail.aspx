<%@ Page Language="C#" AutoEventWireup="true" CodeFile="InvalidUrldetail.aspx.cs" Inherits="GipEditML_InvalidUrldetail" enableEventValidation ="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>生日卡管理</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link rel="stylesheet" href="/inc/setstyle.css" />
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <script language="javascript" type="text/javascript" src="../js/jquery.js"></script>

</head>
<body>
    <form id="form1" runat="server" enableviewstate="False">
    <div>
        <div id="FuncName">
            <h1>
                功能管理／連結失效資料管理</h1>
            <div id="Nav">
                <a href="#" onclick="javascript:history.back();" title="新增">回前頁</a> 
            </div>
            <div id="ClearFloat">
            </div>
        </div>
        <div id="FormName">
        </div>
        <table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
            <tr>
                <td>
                <asp:Button ID="btexcel" runat="server" Text="匯出Excel" OnClick ="ExportToExcel" />
                    <asp:Button ID="btclose" runat="server" Text="將所有資料下架設為不公開(如果為新聞資料URL清空但不下架)" OnClientClick="return confirm('確定執行嗎?')" OnClick="Close_all" />
                </td>
            </tr>
            <tr>
                <td width="95%" colspan="2" height="230" valign="top">
                    <div align="center">
                        <font size="2" color="rgb(63,142,186)">第 <font color="#FF0000">
                            <%=pl.PageIndex+1 %>/<%=pl.TotalPages %></font>&nbsp;頁&nbsp;|&nbsp;共 <font color="#FF0000">
                                <%=pl.TotalCount %></font>&nbsp;筆&nbsp;|&nbsp;跳至第
                            <select size="1" style="color: #FF0000" onchange="page(this.value);">
                                <%=pl.PageOptions %>
                            </select>
                            &nbsp;頁&nbsp;|&nbsp;每頁筆數:&nbsp; </font>
                        <select size="1" style="color: #FF0000" onchange="pagesize(this.value);">
                            <%=pl.PageSizeOptions %>
                        </select>
                        <asp:Repeater ID="rptList" runat="server" OnItemDataBound="R1_ItemDataBound"  >
                            <HeaderTemplate>
                                <table cellspacing="0" id="ListTable" >
                                    <tr>
                                        <th align="left" runat="server" Visible="<%# !IsDataRemoved %>" >
                                            <asp:CheckBox ID="CheckSelectAll"  Text="全選" runat="server" />
                                        </th>
                                        <th>
                                            文件代碼
                                        </th>
                                        <th>
                                            預覽
                                        </th>
                                        <th>
                                            狀態
                                        </th>
										<th >
                                            來源
                                        </th>
                                        <th >
                                            單元名稱
                                        </th>
                                        <th >
                                            類型
                                        </th>                                        
                                        <th >
                                            標題
                                        </th>
                                        <th >
                                            Return Code
                                        </th>
                                    </tr>
                            </HeaderTemplate>
                            <ItemTemplate>
                                <tr>
                                    <td class="eTableContent" runat="server" Visible="<%# !IsDataRemoved %>" >
                                        <asp:CheckBox ID="CheckSelect" runat="server" />
                                        <asp:TextBox id="tbxid" runat="server" Text='<%#Eval("ID") %>' style="display:none" /> 
                                        <asp:TextBox id="tbxacid" runat="server" Text='<%#Eval("ArticleId") %>' style="display:none" /> 
                                    </td>
                                    <td><%#Eval("ArticleId")%></td>
                                    <td class="eTableContent">
                                    <asp:Literal ID="lview" runat="server"></asp:Literal>
                                    </td>
                                    <td class="eTableContent" style="width:60px; height:20px">
                                        <asp:Label ID="lstatus" runat="server"></asp:Label>
                                    </td>
                                    <td class="eTableContent">
                                        <asp:Label ID="lsource" runat="server"></asp:Label>
                                    </td>
                                    <td class="eTableContent">
                                        <asp:Label ID="lunit" runat="server"></asp:Label>
                                    </td>
                                    <td class="eTableContent">
                                        <asp:Label ID="linkType" runat="server"></asp:Label>
                                    </td>                                    
                                    <td class="eTableContent" >
                                        <asp:Label ID="ltitle" runat="server"></asp:Label>
                                    </td>
                                    <td class="eTableContent" >
                                        <%#Eval("Result")%>
                                    </td>
                                </tr>
                            </ItemTemplate>
                            <FooterTemplate>
                                </table></FooterTemplate>
                        </asp:Repeater>
                    </div>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:Button ID="removeHead" runat="server" Text="刪除此次檢查記錄" OnClientClick="return confirm('確定要刪除此次檢查記錄嗎?')" OnClick="RemoveHead" />
                    <asp:Button ID="btremove" runat="server" Text="這是有效的連結，移出列表" OnClientClick="return confirm('確定為有效連結？將移出列表')" OnClick="Remove_select" />
                </td>
            </tr>
        </table>
    </div>    
    <div id="gvdiv" runat="server" style="display: none">
        <%//匯出Excel檔用 %>
        <asp:GridView ID="gridview1" runat="server" OnRowDataBound ="GV_DataBound" AutoGenerateColumns="false">
            <Columns>
                <asp:TemplateField ItemStyle-HorizontalAlign="Left" >
                    <ItemTemplate>
                        <%# Container.DataItemIndex + 1 %></ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="ArticleId" HeaderText="編號" ReadOnly="True" SortExpression="ArticleId" ItemStyle-HorizontalAlign="Left" />
                <asp:TemplateField HeaderText="編修日期" ItemStyle-HorizontalAlign="Left">
                    <ItemTemplate>
                        <asp:Label ID="removeDate" runat="server"></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="來源">
                    <ItemTemplate>
                        <asp:Label ID="source" runat="server"></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="標題名稱">
                    <ItemTemplate>
                        <asp:Label ID="title" runat="server"></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="類型">
                    <ItemTemplate>
                        <asp:Label ID="linkType" runat="server"></asp:Label></ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="網址">
                    <ItemTemplate>
                        <asp:Label ID="url" runat="server"></asp:Label></ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
    </div>
    </form>
</body>

<script type="text/javascript">
    function page(p) {
        window.location.href = 'InvalidUrldetail.aspx?id=<%=sid%>&pagesize=<%=pl.PageSize%>&page=' + (p - 1);
    }
    function pagesize(ps) {
        window.location.href = 'InvalidUrldetail.aspx?id=<%=sid%>&pagesize=' + ps;
    }
</script>
    <script type="text/javascript">
        var rptListControl = document.getElementById('<%= rptList.ClientID %>');
        $('input:checkbox[id$=CheckSelectAll]', rptListControl).click(function(e) {
            if (this.checked) {
                $('input:checkbox[id$=CheckSelect]', rptListControl).attr('checked', true);
            }
            else {
                $('input:checkbox[id$=CheckSelect]', rptListControl).removeAttr('checked');
            }
        });

        $('input:checkbox[id$=CheckSelect]', rptListControl).click(function(e) {
            //To uncheck the header checkbox when there are no selected checkboxes in itemtemplate
            if ($('input:checkbox[id$=CheckSelect]:checked', rptListControl).length == 0) {
                $('input:checkbox[id$=CheckSelectAll]', rptListControl).removeAttr('checked');
            }
            //To check the header checkbox when there are all selected checkboxes in itemtemplate
            else if ($('input:checkbox[id$=CheckSelect]:checked', rptListControl).length == $('input:checkbox[id$=CheckSelect]', rptListControl).length) {
                $('input:checkbox[id$=CheckSelectAll]', rptListControl).attr('checked', true);
            }
        });
   </script>
</html>
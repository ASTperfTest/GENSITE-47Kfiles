<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PublishEdit.aspx.cs" Inherits="PublishEdit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>�A���Ͳ��@���޲z</title>
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <link type="text/css" rel="stylesheet" href="../css/setstyle.css" />

    <script type="text/javascript" src="http://ajax.microsoft.com/ajax/jquery/jquery-1.4.2.min.js"></script>

</head>
<body>
    <form id="form1" runat="server" enableviewstate="False">
    <div>
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
                <td width="50%" align="left" nowrap="nowrap" class="FormName">
                    �A���Ͳ��@���޲z&nbsp;<font size="2"><asp:Literal id="titles" runat="server"></asp:Literal></font>
                </td>
                <td width="50%" class="FormLink" align="right">
                    <a href="Publish.aspx?id=<%=pid%>" title="�^�e��">�^�e��</a>
                </td>
            </tr>
            <tr>
                <td width="100%" colspan="2">
                    <hr />
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" width="80%" height="230" valign="top">
                    <table border="0" width="100%">
                        <tr>
                            <td width="30%">
                                <%if (topCat == "E")
                                  { %>
                                <fieldset style="width: 300px">
                                    <legend style="font-size: 12px"><b>�s�W�귽�����p��</b></legend>
                                    <table style="font-size: 12px">
                                        <tr>
                                            <td align="left">
                                                &nbsp;&nbsp;���D�G
                                                <input type="text" name="Resource_Title" id="Resource_Title" maxlength="30" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                &nbsp;&nbsp;���}�G
                                                <input type="text" name="Resource_Url" id="Resource_Url" maxlength="100" />
                                                &nbsp;&nbsp;<input type="button" value="�e�X" style="font-size: 12px" onclick="addURL()" />
                                            </td>
                                        </tr>
                                    </table>
                                </fieldset>
                                <%}%>
                            </td>
                            <td width="70%" align="left">
                                <font size="2" color="rgb(63,142,186)">�� <font color="#FF0000">
                                    <%=pl.PageIndex+1 %>/<%=pl.TotalPages %></font>&nbsp;��&nbsp;|&nbsp;�@ <font color="#FF0000">
                                        <%=pl.TotalCount %></font>&nbsp;��&nbsp;|&nbsp;���ܲ�
                                    <select size="1" style="color: #FF0000" onchange="page(this.value);">
                                        <%=pl.PageOptions %>
                                    </select>
                                    &nbsp;��&nbsp;|&nbsp;�C������:&nbsp; </font>
                                <select size="1" style="color: #FF0000" onchange="pagesize(this.value);">
                                    <%=pl.PageSizeOptions %>
                                </select>
                            </td>
                        </tr>
                    </table>
                    <asp:Repeater ID="rptList" runat="server">
                        <HeaderTemplate>
                            <table width="100%" id="ListTable">
                                <tr align="left">
                                    <th width="7%">
                                        <input type="hidden" name="selected" id="selected" value="0" />
                                        <input type="button" class="cbutton" onclick="selectAll()" value="����/������">
                                    </th>
                                    <% if (topCat == "E")
                                       { %>
                                    <th>�w��</th>
                                    <th>���D</th>
                                    <th>URL</th>
                                    <th>����</th>
                                    <th>�s�פ��</th>
									<% }
                                       else if (topCat == "F")
                                       { %>
									<th>�d���H</th>
                                    <th>�d�����</th>
                                    <th>�d�����e</th>
                                    <th>�s�פ��</th>
                                    <% } else { %>
                                    <th>�w��</th>
                                    <th>��Ʈw�ӷ�</th>
                                    <th>�椸�W��</th>
                                    <th>��Ƽ��D</th>
                                    <th>����</th>
                                    <th>�s�פ��</th>
                                    <% } %>
                                </tr>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <% if (topCat == "E")
                               { %>
                            <tr>
                                <td>
                                    <div align="center">
                                        <input type="checkbox" name="ckbox" value="<%# Eval("gicuitem")%>">
                                    </div>
                                </td>
                                <td>
                                    <a href="<%# Eval("path")%>" target="_blank">View</a>
                                </td>
                                <td>
                                    <%# Eval("sTitle")%>
                                </td>
                                <td>
                                    <%# Eval("path")%>
                                </td>
                                <td>
                                    <span class="eTableContent">
                                        <input type="text" name='<%# Eval("gicuitem")%>' value="<%# Eval("orderArticle")%>"
                                            maxlength="5" size="5">
                                    </span>
                                &nbsp;</td>
                                <td>
                                    <%# Eval("dEditDate")%>
                                </td>
                            </tr>
							<% }
                               else if (topCat == "F")
                               { %>
                            <tr>
                                <td align="center">
                                    <div align="center">
                                        <input type="checkbox" name="ckbox" value="<%# Eval("gicuitem")%>">
                                    </div>
                                </td>
                                <td>
                                    <%# Eval("iEditor")%>
                                </td>
                                <td>
                                    <%# Eval("xpostdate")%>
                                </td>
                                <td>
                                    <%# Eval("xBody")%>
                                </td>
                                <td>
                                    <%# Eval("dEditDate")%>
                                </td>
                            </tr>
                            <% }
                               else
                               { %>
                            <tr>
                                <td>
                                    <div align="center">
                                        <input type="checkbox" name="ckbox" value="<%# Eval("gicuitem")%>">
                                    </div>
                                </td>
                                <td>
                                    <a href='<%# Eval("path")%>' target="_blank">
                                        View</a>
                                </td>
                                <td>
                                    <%# ((int)Eval("CtRootId") == 1 ? "�J�f��" : ((int)Eval("CtRootId") == 2 ? "�D�D�]" : ((int)Eval("CtRootId") == 3 ? "���Ѯa" : ((int)Eval("CtRootId") == 4 ? "���Ѯw" : "&nbsp;"))))%>
                                </td>
                                <td>
                                    <%# getUnit((int)Eval("ArticleId"), (int)Eval("CtRootId"), (string)Eval("CtUnitId"))%>
                                </td>
                                <td>
                                    <%# Eval("sTitle")%>
                                </td>
                                <td>
                                    <span class="eTableContent">
                                        <input type="text" name='<%# Eval("gicuitem")%>' value="<%# Eval("orderArticle")%>"
                                            maxlength="5" size="5">
                                    </span>
                                </td>
                                <td>
                                    <%# Eval("dEditDate")%>
                                </td>
                            </tr>
                            <% } %>
                        </ItemTemplate>
                        <FooterTemplate>
                            </table></FooterTemplate>
                    </asp:Repeater>
                    <input type="hidden" name="action" id="action" value="" />
                    <input type="button" class="cbutton" onclick="del()" value="�R�����" />
                    <% if (topCat != "F")
                       { %>
                    <input type="button" class="cbutton" onclick="edit()" value="�s�צs��" />
                    <% } %>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>

<script type="text/javascript">
    function page(p) {
        window.location.href = 'PublishEdit.aspx?topCat=<%=topCat%>&pid=<%=pid%>&id=<%=id%>&pagesize=<%=pl.PageSize%>&page=' + (p - 1);
    }
    function pagesize(ps) {
        window.location.href = 'PublishEdit.aspx?topCat=<%=topCat%>&pid=<%=pid%>&id=<%=id%>&pagesize=' + ps;
    }
    function selectAll() {
        $(":checkbox").each(function() {
            $(this).attr("checked", $("#selected").val() == "0" ? true : false);
        })
        $("#selected").val($("#selected").val() == "0" ? "1" : "0");
    }
    function del() {
        $("#action").val("del");
        form1.submit();
    }
    function edit() {
        $("#action").val("edit");
        form1.submit();
    }
    function addURL() {
        $("#action").val("addURL");
        if ($("#Resource_Title").val() == "") {
            alert('���D���i���ť�');
            form1.Resource_Title.focus();
            return;
        }
        if ($("#Resource_Url").val() == "") {
            alert('���}���i���ť�');
            form1.Resource_Url.focus();
            return;
        }
        form1.submit();
    }
</script>

</html>

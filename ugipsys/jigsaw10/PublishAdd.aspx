<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PublishAdd.aspx.cs" Inherits="PublishAdd" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>農漁生產作物管理</title>
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
                    農漁生產作物管理&nbsp;<font size="2"><asp:Literal id="titles" runat="server"></asp:Literal></font>
                </td>
                <td width="50%" class="FormLink" align="right">
                    <a href="PublishQuery.aspx?pid=<%=pid%>&id=<%=id%>&topCat=" title="設定單元內容文章">設定單元內容文章</a>
                    <a href="PublishQuery.aspx?pid=<%=pid%>&id=<%=id%>&topCat=<%=topCat%>" title="回前頁">回前頁</a>
                </td>
            </tr>
            <tr>
                <td width="100%" colspan="2">
                    <hr />
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" width="80%" height="230" valign="top">
                    <p align="center">
                        <font size="2" color="rgb(63,142,186)">第 <font color="#FF0000">
                            <%=pl.PageIndex+1 %>/<%=pl.TotalPages %></font>&nbsp;頁&nbsp;|&nbsp;共 <font color="#FF0000">
                                <%=pl.TotalCount %></font>&nbsp;筆&nbsp;|&nbsp;跳至第
                            <select size="1" style="color: #FF0000" onchange="page(this.value);">
                                <%=pl.PageOptions %>
                            </select>
                            &nbsp;頁
                            <% if (pl.HasPreviousPage)
                               { %>
                            |<a href="javascript:page(<%=pl.PageIndex %>)">上一頁</a>
                            <% } %>
                            <% if (pl.HasNextPage)
                               { %>
                            |<a href="javascript:page(<%=pl.PageIndex+2 %>)">下一頁</a>
                            <% } %>
                            |&nbsp;每頁筆數:&nbsp; </font>
                        <select size="1" style="color: #FF0000" onchange="pagesize(this.value);">
                            <%=pl.PageSizeOptions %>
                        </select>
                    </p>
                    <table cellspacing="1" cellpadding="0" width="100%" align="center" border="0">
                        <tbody>
                            <tr>
                                <td valign="top" width="95%" colspan="2" height="230">
                                    <center>
                                        <asp:Repeater ID="rptList" runat="server">
                                            <HeaderTemplate>
                                                <table width="100%" id="ListTable">
                                                    <tbody>
                                                        <tr align="left">
                                                            <th width="7%">
                                                                <input type="hidden" name="selected" id="selected" value="0" />
                                                                <input type="button" class="cbutton" onclick="selectAll()" value="全選/全不選">
                                                            </th>
                                                            <th>
                                                                預覽
                                                            </th>
                                                            <th>
                                                                資料庫來源
                                                            </th>
                                                            <th>
                                                                節點名稱
                                                            </th>
                                                            <th>
                                                                資料標題
                                                            </th>
                                                            <th>
                                                                帳號
                                                            </th>
                                                            <th>
                                                                單位
                                                            </th>
                                                            <th>
                                                                編修日期
                                                            </th>
                                                        </tr>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <tr>
                                                    <td>
                                                        <div align="center">
                                                            <%# getCheckbox((int)Eval("iCUItem"), id)%>
                                                        </div>
                                                    </td>
                                                    <td>
                                                        <%# getAnchor(CtRootId, (int)Eval("iCUItem"), (int)Eval("categoryId"))%>
                                                    </td>
                                                    <td>
                                                        <%= (CtRootId == "1" ? "入口網" : (CtRootId == "2" ? "主題館" : (CtRootId == "3" ? "知識家" : (CtRootId == "4" ? "知識庫" : "&nbsp;"))))%>
                                                    </td>
                                                    <td>
                                                        <%# Eval("CatName")%>
                                                    </td>
                                                    <td>
                                                        <%# Eval("sTitle")%>
                                                    </td>
                                                    <td>
                                                        <%# Eval("UserName")%>
                                                    </td>
                                                    <td>
                                                        <%# Eval("deptName")%>
                                                    </td>
                                                    <td>
                                                        <%# Eval("dEditDate")%>
                                                    </td>
                                                </tr>
                                            </ItemTemplate>
                                            <FooterTemplate>
                                                </tbody></table>
                                            </FooterTemplate>
                                        </asp:Repeater>
                                    </center>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <input type="hidden" name="CtRootId" value="<%= Server.UrlEncode(CtRootId) %>" />
                    <input type="submit" value="編修存檔" class="cbutton" />
                    <input type="button" value="回上層" class="cbutton" onclick="location.href='Publish.aspx?id=<%=pid%>'" />
                    <input type="reset" value="重設" class="cbutton" />
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>

<script type="text/javascript">
    function page(p) {
        window.location.href = 'PublishAdd.aspx?pid=<%=pid%>&id=<%=id%>&topCat=<%=topCat%>&CtRootId=<%=CtRootId%>&CtNodeName=<%=Server.UrlEncode(CtNodeName)%>&sTitle=<%=Server.UrlEncode(sTitle)%>&Status=<%=Status%>&pagesize=<%=pl.PageSize%>&page=' + (p - 1);
    }
    function pagesize(ps) {
        window.location.href = 'PublishAdd.aspx?pid=<%=pid%>&id=<%=id%>&topCat=<%=topCat%>&CtRootId=<%=CtRootId%>&CtNodeName=<%=Server.UrlEncode(CtNodeName)%>&sTitle=<%=Server.UrlEncode(sTitle)%>&Status=<%=Status%>&pagesize=' + ps;
    }
    function selectAll() {
        $(":checkbox").each(function() {
            $(this).attr("checked", $("#selected").val() == "0" ? true : false);
        })
        $("#selected").val($("#selected").val() == "0" ? "1" : "0");
    }
</script>

</html>

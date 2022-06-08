<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PublishEdit.aspx.cs" Inherits="PublishEdit" %>

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
                    <a href="Publish.aspx?id=<%=pid%>" title="回前頁">回前頁</a>
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
                                    <legend style="font-size: 12px"><b>新增資源推薦聯結</b></legend>
                                    <table style="font-size: 12px">
                                        <tr>
                                            <td align="left">
                                                &nbsp;&nbsp;標題：
                                                <input type="text" name="Resource_Title" id="Resource_Title" maxlength="30" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                &nbsp;&nbsp;網址：
                                                <input type="text" name="Resource_Url" id="Resource_Url" maxlength="100" />
                                                &nbsp;&nbsp;<input type="button" value="送出" style="font-size: 12px" onclick="addURL()" />
                                            </td>
                                        </tr>
                                    </table>
                                </fieldset>
                                <%}%>
                            </td>
                            <td width="70%" align="left">
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
                            </td>
                        </tr>
                    </table>
                    <asp:Repeater ID="rptList" runat="server">
                        <HeaderTemplate>
                            <table width="100%" id="ListTable">
                                <tr align="left">
                                    <th width="7%">
                                        <input type="hidden" name="selected" id="selected" value="0" />
                                        <input type="button" class="cbutton" onclick="selectAll()" value="全選/全不選">
                                    </th>
                                    <% if (topCat == "E")
                                       { %>
                                    <th>預覽</th>
                                    <th>標題</th>
                                    <th>URL</th>
                                    <th>順序</th>
                                    <th>編修日期</th>
									<% }
                                       else if (topCat == "F")
                                       { %>
									<th>留言人</th>
                                    <th>留言日期</th>
                                    <th>留言內容</th>
                                    <th>編修日期</th>
                                    <% } else { %>
                                    <th>預覽</th>
                                    <th>資料庫來源</th>
                                    <th>單元名稱</th>
                                    <th>資料標題</th>
                                    <th>順序</th>
                                    <th>編修日期</th>
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
                                    <%# ((int)Eval("CtRootId") == 1 ? "入口網" : ((int)Eval("CtRootId") == 2 ? "主題館" : ((int)Eval("CtRootId") == 3 ? "知識家" : ((int)Eval("CtRootId") == 4 ? "知識庫" : "&nbsp;"))))%>
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
                    <input type="button" class="cbutton" onclick="del()" value="刪除選擇" />
                    <% if (topCat != "F")
                       { %>
                    <input type="button" class="cbutton" onclick="edit()" value="編修存檔" />
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
            alert('標題不可為空白');
            form1.Resource_Title.focus();
            return;
        }
        if ($("#Resource_Url").val() == "") {
            alert('網址不可為空白');
            form1.Resource_Url.focus();
            return;
        }
        form1.submit();
    }
</script>

</html>

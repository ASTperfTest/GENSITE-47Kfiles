<%@Page Language="C#" AutoEventWireup="true" CodeFile="Index.aspx.cs" Inherits="jigsaw10_index" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>農作物地圖管理</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link rel="stylesheet" href="/inc/setstyle.css" />
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
</head>
<body>
    <form id="form1" runat="server" enableviewstate="False">
    <div>
        <div id="FuncName">
            <h1>
                農漁生產地圖管理<font size="2">【作物清單】</font></h1>
            <div id="Nav">
				<a href="PhoneticNotation.aspx" title="罕見字注音管理">罕見字注音管理</a> 
                <a href="EditCrop.aspx" title="新增">新增作物</a> 
				<a href="Search.aspx" title="查詢">作物查詢</a>
            </div>
            <div id="ClearFloat">
            </div>
        </div>
        <div id="FormName">
        </div>
        <table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
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
                        <asp:Repeater ID="rptList" runat="server">
                            <HeaderTemplate>
                                <table cellspacing="0" id="ListTable">
                                    <tr>
                                        <th>
                                            主題作物名稱
                                        </th>
										<th style='width:50px'>
                                            預覽
                                        </th>
                                        <th>
                                            類型
                                        </th>
                                        <th>
                                            是否公開
                                        </th>
                                        <th>
                                            設定作物內容
                                        </th>
                                    </tr>
                            </HeaderTemplate>
                            <ItemTemplate>
                                <tr>
                                    <td class="eTableContent">
                                        <a href="<%# (char)Eval("RSS") == '0' ? "EditCrop.aspx" : "EditFish.aspx"%>?id=<%# Eval("iCUItem")%>&returnUrl=<%=Server.UrlEncode(Request.Url.PathAndQuery)%>">
                                            <%# Eval("sTitle")%></a>
                                    </td>
									<td align="center" class="eTableContent">
                                        <a target="_blank" href="<%# StrFunc.GetWebSetting("WebURL", "appSettings").Replace("\\", "/") %>/jigsaw2010/detail.aspx?item=<%# Eval("iCUItem")%>">
                                            view</a>
                                    </td>
                                    <td align="center" class="eTableContent">
                                        <%# Eval("Categories")%>
                                    </td>
                                    <td align="center" class="eTableContent">
                                        <%# (char)Eval("fCTUPublic") == 'Y' ? "公開" : "不公開"%>
                                    </td>
                                    <td align="center" class="eTableContent">
                                        <input name="button" type="button" class="cbutton" onclick="location.href='Publish.aspx?id=<%# Eval("iCUItem")%>&returnUrl=<%=Server.UrlEncode(Request.Url.PathAndQuery)%>'"
                                            value="設定">
                                    </td>
                                </tr>
                            </ItemTemplate>
                            <FooterTemplate>
                                </table></FooterTemplate>
                        </asp:Repeater>
                    </div>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>

<script type="text/javascript">
    function page(p) {
        window.location.href = 'Index.aspx?sTitle=<%=Server.UrlEncode(titles)%>&Status=<%=status%>&Types=<%=types%>&pagesize=<%=pl.PageSize%>&page=' + (p - 1);
    }
    function pagesize(ps) {
        window.location.href = 'Index.aspx?sTitle=<%=Server.UrlEncode(titles)%>&Status=<%=status%>&Types=<%=types%>&pagesize=' + ps;
    }
</script>

</html>

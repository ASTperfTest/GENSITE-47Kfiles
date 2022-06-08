<%@ Page Language="C#" AutoEventWireup="true" CodeFile="InvalidUrlHead.aspx.cs" Inherits="GipEditML_InvalidUrlHead" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>連結失效管理</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link rel="stylesheet" href="/inc/setstyle.css" />
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <script language="javascript" type="text/javascript" src="../js/jquery.js"></script>

    <script type="text/javascript">
        function NewAjax(id, url, Eventjs, dataArr) {
            $.ajax({
                type: 'post',
                url: url,
                dataType: "text",
                data: dataArr,
                cache: false,
                success: function(html) {
                    if (id != "") {
                        $("#" + id).html(html);
                    }
                    if (typeof Eventjs == "function") {
                        Eventjs();
                    }
                    else if (typeof Eventjs == "string") {
                        eval(Eventjs);
                    }
                }
            });
        }
    </script>

    <script type="text/javascript">

        function CheckURL() {
            window.open("<%=callerURL %>&newact=Y&Id=0");
            window.location.reload();
        }


        function react(sid) {
            window.open("<%=callerURL %>&newact=N&Id=" + sid);
            window.location.reload();
        }
        
    </script>
</head>
<body>
    <form id="form1" runat="server" enableviewstate="False">
    <div>
        <div id="FuncName">
            <h1>
                功能管理／連結失效資料管理</h1>
            <div id="Nav">
                <a href="javascript:CheckURL()" id="checkurl">啟動檢查</a> 
            </div>                
        </div>
        <div id="FormName">
        </div>
        <table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
                <tr>
                    <td width="95%" colspan="2" height="230" valign="top">
                    <div>
                        <asp:Repeater ID="rptList" runat="server" OnItemCommand="Repeater_OnItemCommand" OnItemDataBound="R1_ItemDataBound">
                            <HeaderTemplate>
                                <table cellspacing="0" id="ListTable">
                                    <tr>
                                        <th>
                                            檢查時間
                                        </th>
                                        <th>
                                            檢查進度
                                        </th>
                                        <th>
                                            移除時間
                                        </th>
                                        <th>
                                            文章列表
                                        </th>
                                    </tr>
                            </HeaderTemplate>
                            <ItemTemplate>
                                <tr>
                                    <td class="eTableContent">
                                        <%#String.Format("{0:d}", Eval("runDate"))%>
                                    </td>
                                    <td class="eTableContent">
                                    <asp:Literal ID="lblProcess" runat="server"></asp:Literal>
                                    </td>
                                    <td class="eTableContent">
                                        <%#String.Format("{0:d}", Eval("removeDate"))%>
                                    </td>
                                    <td class="eTableContent">
                                        <asp:LinkButton ID="lnkDetail" runat="server" Text="文章列表" CommandArgument='<%#Eval("ID") %>'
                                            CommandName="Select"></asp:LinkButton>
                                    </td>
                                </tr>
                            </ItemTemplate>
                            <FooterTemplate>
                                </table>
                            </FooterTemplate>
                        </asp:Repeater>
                        </div>
                    </td>
                </tr>
            </table>
    </div>
    </form>
</body>
</html>

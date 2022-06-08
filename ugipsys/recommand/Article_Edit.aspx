<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Article_Edit.aspx.cs" Inherits="Article_Edit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>報馬仔文章審核</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <%-- jQuery UI - MultiSelect --%>
    <link rel="stylesheet" href="../js/multiselect/css/common.css" type="text/css" />
    <link type="text/css" href="../js/multiselect/css/ui.multiselect.css" rel="stylesheet" />
    <link type="text/css" href=" rel="stylesheet" />
    <link type="text/css" rel="stylesheet" href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.8.10/themes/ui-lightness/jquery-ui.css" />
    <script type="text/javascript" src="../js/multiselect/js/jquery-1.6.1.min.js"></script>
    <script type="text/javascript" src="../js/multiselect/js/jquery-ui-1.8.14.custom.min.js"></script>
    <script type="text/javascript" src="../js/multiselect/js/ui.multiselect.js"></script>
    <script language="javascript" type="text/javascript">
        $(function () {
            $(".multiselect").multiselect();
        });
    </script>
    <%-- jQuery UI - MultiSelect --%>

</head>
<body>
    <form id="form1" runat="server" enableviewstate="False">
    <div>
        <div id="FuncName">
            <h1>
                報馬仔文章管理 / 報馬仔文章審核 <font size="2">【編修專區】</font></h1>
            <div id="Nav">
            </div>
            <div id="ClearFloat">
            </div>
        </div>
        <div id="FormName">
        </div>
        <table id="ListTable">
            <tr>
                <th style="width:20%;"><font color="red">*</font>推薦好文標題：</th>
                <td class="eTableContent"><asp:TextBox runat="server" ID="txtTitle" MaxLength="50" AutoCompleteType="Disabled" Width="85%" /></td>
            </tr>
            <tr>
                <th><font color="red">*</font>推薦好文網址：</th>
                <td class="eTableContent"><asp:TextBox runat="server" ID="txtURL" MaxLength="200" AutoCompleteType="Disabled" Width="85%" /></td>
            </tr>
            <tr>
                <th>推薦原因：</th>
                <td class="eTableContent"><asp:TextBox runat="server" ID="txtContent" Width="85%" TextMode="MultiLine" Rows="7"  /></td>
            </tr>
            <tr>
            <tr>
                <th><font color="red">*</font>資料出處(請輸入資料來源資訊。<br />例如：wikipedia、xxx部落客 等等)：</th>
                <td class="eTableContent"><asp:TextBox id="txtSource" runat="server" Width="85%" MaxLength="50" AutoCompleteType="Disabled"></asp:TextBox></td>
            </tr>
            </tr>
            <tr>
                <th>標籤：</th>
                <td class="eTableContent">點選以下標籤：<br />
                <asp:CheckBoxList ID = "listTAGs" runat = "server" RepeatDirection="Horizontal"  ></asp:CheckBoxList> 
                <%--<asp:ListBox runat="server" ID="listTAGs" Width="500px" cssclass="multiselect" SelectionMode="Multiple"></asp:ListBox><br />--%>
            </td>
            </tr>
            <tr>
                <td  colspan="2" style="text-align: center;">
                    <asp:Button runat="server" ID="btnOK" CssClass="cbutton" Text="編修存檔" OnClick="btnOK_Click" />
                    &nbsp;&nbsp;
                    <input type="button" value="取消並回列表" class="cbutton" onclick="javascript:void(location.href='Index.aspx');" />
                    <%--<input type="button" value="回前頁" class="cbutton" onclick="javascript:void(history.back());" />--%>
                </td>
            </tr>
        </table>
        <script type="text/javascript">
            var errString = '<%=errString %>';
            var saveOK = '<%=saveOK %>';
            if (errString != '') {
                alert(errString);
            }
            if (saveOK == "Y") {
                alert('資料編修完成。');
                window.location.href = 'Index.aspx';
            }
        </script>
    </div>
    </form>
</body>
</html>

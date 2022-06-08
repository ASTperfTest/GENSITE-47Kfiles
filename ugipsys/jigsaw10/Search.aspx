<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Search.aspx.cs" Inherits="Search" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>查詢表單</title>
    <link type="text/css" rel="stylesheet" href="../css/list.css"/>
	<link type="text/css" rel="stylesheet" href="../css/layout.css"/>
	<link type="text/css" rel="stylesheet" href="../css/setstyle.css"/>
</head>
<body>
    <form method="post" action="Index.aspx">
    <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr>
            <td width="50%" class="FormName" align="left">
                農漁作物管理&nbsp;<font size="2">【作物查詢】</font>
            </td>
            <td width="50%" class="FormLink" align="right">
                <a href="Javascript:window.history.back();" title="回前頁">回前頁</a>
            </td>
        </tr>
        <tr>
            <td width="100%" colspan="2">
                <hr />
            </td>
        </tr>
        <tr>
            <td class="Formtext" colspan="2" height="15">
            </td>
        </tr>
        <tr>
            <td align="center" colspan="2" width="80%" height="230" valign="top">
                <center>
                    <table width="100%" id="ListTable">
                        <tr>
                            <th>
                                資料標題：
                            </th>
                            <td class="eTableContent">
                                <label>
                                    <input name="sTitle" id="sTitle" size="30" />
                                    (輸入資料標題)</label>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                狀態：
                            </th>
                            <td class="eTableContent">
                                <span class="FormLink">
                                    <select name="Status" id="Status">
                                        <option value="">ALL</option>
                                        <option value="Y" selected="selected">公開</option>
                                        <option value="N">不公開</option>
                                    </select>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                類型：
                            </th>
                            <td class="eTableContent">
                                <span class="FormLink">
                                    <select name="Types" id="Types">
                                        <option value="" selected="selected">ALL</option>
                                        <option value="0" >農作物</option>
                                        <option value="1">魚種</option>
                                    </select>
                                </span>
                            </td>
                        </tr>
                    </table>
                </center>
                <table border="0" width="100%" cellspacing="0" cellpadding="0">
                    <tr>
                        <td align="center">
                            <input type="submit" value="查詢" class="cbutton" />
                            <input type="reset" value="重　填" class="cbutton" />
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>

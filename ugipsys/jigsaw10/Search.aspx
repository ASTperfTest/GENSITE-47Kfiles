<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Search.aspx.cs" Inherits="Search" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>�d�ߪ��</title>
    <link type="text/css" rel="stylesheet" href="../css/list.css"/>
	<link type="text/css" rel="stylesheet" href="../css/layout.css"/>
	<link type="text/css" rel="stylesheet" href="../css/setstyle.css"/>
</head>
<body>
    <form method="post" action="Index.aspx">
    <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr>
            <td width="50%" class="FormName" align="left">
                �A���@���޲z&nbsp;<font size="2">�i�@���d�ߡj</font>
            </td>
            <td width="50%" class="FormLink" align="right">
                <a href="Javascript:window.history.back();" title="�^�e��">�^�e��</a>
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
                                ��Ƽ��D�G
                            </th>
                            <td class="eTableContent">
                                <label>
                                    <input name="sTitle" id="sTitle" size="30" />
                                    (��J��Ƽ��D)</label>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                ���A�G
                            </th>
                            <td class="eTableContent">
                                <span class="FormLink">
                                    <select name="Status" id="Status">
                                        <option value="">ALL</option>
                                        <option value="Y" selected="selected">���}</option>
                                        <option value="N">�����}</option>
                                    </select>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                �����G
                            </th>
                            <td class="eTableContent">
                                <span class="FormLink">
                                    <select name="Types" id="Types">
                                        <option value="" selected="selected">ALL</option>
                                        <option value="0" >�A�@��</option>
                                        <option value="1">����</option>
                                    </select>
                                </span>
                            </td>
                        </tr>
                    </table>
                </center>
                <table border="0" width="100%" cellspacing="0" cellpadding="0">
                    <tr>
                        <td align="center">
                            <input type="submit" value="�d��" class="cbutton" />
                            <input type="reset" value="���@��" class="cbutton" />
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>

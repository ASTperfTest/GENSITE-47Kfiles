<%@ Page Language="C#" AutoEventWireup="true" CodeFile="EditCrop.aspx.cs" Inherits="EditCrop" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>�A�@�����@</title>
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <link type="text/css" rel="stylesheet" href="../css/setstyle.css" />
	<script type="text/javascript" src="http://ajax.microsoft.com/ajax/jquery/jquery-1.4.2.min.js"></script>
</head>
<body>
    <form id="form1" enctype="multipart/form-data" runat="server" 
    enableviewstate="False">
    <div>
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr>
                <td width="50%" class="FormName" align="left">
                    �A���@���޲z&nbsp;<font size="2">�i�A�@�����@�j</font>
                </td>
                <td width="50%" class="FormLink" align="right">
                    <!--Added By Leo    2011-06-20      �W�[�W�ǽs��@��    Start   -->
                    <a  title="�W�ǽs��@���Ӥ�" onclick="javascript:void(window.open('Add_Image.aspx?item=' + <%=Request.QueryString["id"]%> ,'','toolbar=no,menubar=no,scrollbars=yes,location=no,status=no,width=500px,height=500px'));">�W�ǽs��@���Ӥ�</a>
                    <!--Added By Leo    2011-06-20      �W�[�W�ǽs��@��     End    -->
                    <a href="<%=returnUrl%>" title="�^�e��">�^�e��</a>
                </td>
            </tr>
            <tr>
                <td width="100%" colspan="2">
                    <hr />
                    <font color="red">*</font><font size="2"> �N����J����</font><br />
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" width="80%" height="230" valign="top">
                    <table id="ListTable">
                        <tr>
                            <th>
                                <font color="red">*</font> �����G
                            </th>
                            <td class="eTableContent">
                                <select name="CropOrFish" id="CropOrFish" runat="server" onchange="CropOrFish_onchange(this.value);">
                                    <option value="0" selected="selected">�A�@��</option>
                                    <option value="1">����</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                <font color="red">*</font> �@���W�١G
                            </th>
                            <td class="eTableContent">
                                <input name="sTitle" id="sTitle" size="50" value="(�@���W��)" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <th>
                                ���ɡG
                            </th>
                            <td class="eTableContent">
                                <a href="<%=xImgURL%>" target="_blank">
                                    <img alt="" src="<%=xImgURL%>" width="80" border="0" /></a>
                                <input type="file" name="xImgFile" id="xImgFile" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <th>
                                �ƧǡG
                            </th>
                            <td class="eTableContent">
                                �K�G<input name="sortOrderSpring" id="sortOrderSpring" size="1" value="0" runat="server" />&nbsp;�L�G<input
                                    name="sortOrderSummer" id="sortOrderSummer" size="1" value="0" runat="server" />&nbsp;��G<input
                                        name="sortOrderAutumn" id="sortOrderAutumn" size="1" value="0" runat="server" />&nbsp;�V�G<input
                                            name="sortOrderWinter" id="sortOrderWinter" size="1" value="0" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <th>
                                �O�_���}�G
                            </th>
                            <td class="eTableContent">
                                <select name="fCTUPublic" id="fCTUPublic" runat="server">
                                    <option value="Y" selected="selected">���}</option>
                                    <option value="N">�����}</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                <font color="red">*</font> �@�������G
                            </th>
                            <td class="eTableContent">
                                <select name="type" id="type" runat="server">
                                    <option value="">== �п�� ==</option>
                                    <option value="0">���G</option>
                                    <option value="1">����</option>
                                    <option value="2">��c</option>
                                    <option value="3">��³�S�@</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                �@��²���G
                            </th>
                            <td class="eTableContent">
                                <textarea name="description" id="description" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                �O�W�G
                            </th>
                            <td class="eTableContent">
                                <textarea name="alias" id="alias" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                <font color="red">*</font> �����G
                            </th>
                            <td class="eTableContent">
                                <textarea name="season" id="season" cols="50" rows="7" runat="server" value="(����)"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                <font color="red">*</font> ���a�G
                            </th>
                            <td class="eTableContent">
                                <textarea name="origin" id="origin" cols="50" rows="7" runat="server" value="(���a)"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                �S�ʡG
                            </th>
                            <td class="eTableContent">
                                <textarea name="feature" id="feature" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                ��i���ȡG
                            </th>
                            <td class="eTableContent">
                                <textarea name="nutritionValue" id="nutritionValue" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                �D���k�G
                            </th>
                            <td class="eTableContent">
                                <textarea name="selectionMethod" id="selectionMethod" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                ��i�����G
                            </th>
                            <td class="eTableContent">
                                <textarea name="nutrient" id="nutrient" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                �Ʋz�G
                            </th>
                            <td class="eTableContent">
                                <textarea name="cuisine" id="cuisine" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                �~�ءG
                            </th>
                            <td class="eTableContent">
                                <textarea name="variety" id="variety" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                �Ƶ��G
                            </th>
                            <td class="eTableContent">
                                <textarea name="note" id="note" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                �p���Z�G
                            </th>
                            <td class="eTableContent">
                                <textarea name="tips" id="tips" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                    </table>
                    <table border="0" width="100%" cellspacing="0" cellpadding="0">
                        <tr>
                            <td align="center">
                                <input type="hidden" name="action_equals_Del" id="action_equals_Del" value="0" />
                                <input type="button" value="�s�צs��" class="cbutton" onclick="edit()" />
                                <% if (id != 0)
                                   { %>
                                <input type="button" value="�R�@��" class="cbutton" onclick="del()" />
                                <%} %>
                                <input type="reset" value="���@��" class="cbutton" />
                                <input type="button" value="�^�e��" class="cbutton" onclick="location.href='<%=returnUrl%>'" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>

<script type="text/javascript">
    function CropOrFish_onchange(s) {
        if (s == "1")
            window.location.href = "EditFish.aspx";
    }
    function del() {
        $("#action_equals_Del").val("1");
        form1.submit();
    }
    function edit() {
        $("#action_equals_Del").val("0");
        form1.submit();
    }
</script>

</html>

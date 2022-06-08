<%@ Page Language="C#" AutoEventWireup="true" CodeFile="EditFish.aspx.cs" Inherits="EditFish" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>魚種維護</title>
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <link type="text/css" rel="stylesheet" href="../css/setstyle.css" />

    <script type="text/javascript" src="http://ajax.microsoft.com/ajax/jquery/jquery-1.4.2.min.js"></script>

</head>
<body>
    <form id="form1" enctype="multipart/form-data" runat="server" enableviewstate="False">
    <div>
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr>
                <td width="50%" class="FormName" align="left">
                    農漁作物管理&nbsp;<font size="2">【魚種維護】</font>
                </td>
                <td width="50%" class="FormLink" align="right">
                    <!--Added By Leo    2011-06-20      增加上傳編輯作物    Start   -->
                    <a  title="上傳編輯作物照片" onclick="javascript:void(window.open('Add_Image.aspx?item=' + <%=Request.QueryString["id"]%> ,'','toolbar=no,menubar=no,scrollbars=yes,location=no,status=no,width=500px,height=500px'));">上傳編輯作物照片</a>
                    <!--Added By Leo    2011-06-20      增加上傳編輯作物     End    -->
                    <a href="<%=returnUrl%>" title="回前頁">回前頁</a>
                </td>
            </tr>
            <tr>
                <td width="100%" colspan="2">
                    <hr />
                    <font color="red">*</font><font size="2"> 代表必輸入項目</font><br />
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" width="80%" height="230" valign="top">
                    <table id="ListTable">
                        <tr>
                            <th>
                                <font color="red">*</font> 類型：
                            </th>
                            <td class="eTableContent">
                                <select name="CropOrFish" id="CropOrFish" runat="server" onchange="CropOrFish_onchange(this.value);">
                                    <option value="0">農作物</option>
                                    <option value="1" selected="selected">魚種</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                <font color="red">*</font> 中文名稱：
                            </th>
                            <td class="eTableContent">
                                <input name="sTitle" id="sTitle" size="50" value="(中文名稱)" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <th>
                                圖檔：
                            </th>
                            <td class="eTableContent">
                                <a href="<%=xImgURL%>" target="_blank">
                                    <img alt="" src="<%=xImgURL%>" width="80" border="0" /></a>
                                <input type="file" name="xImgFile" id="xImgFile" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <th>
                                排序：
                            </th>
                            <td class="eTableContent">
                                春：<input name="sortOrderSpring" id="sortOrderSpring" size="1" value="0" runat="server" />&nbsp;夏：<input
                                    name="sortOrderSummer" id="sortOrderSummer" size="1" value="0" runat="server" />&nbsp;秋：<input
                                        name="sortOrderAutumn" id="sortOrderAutumn" size="1" value="0" runat="server" />&nbsp;冬：<input
                                            name="sortOrderWinter" id="sortOrderWinter" size="1" value="0" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <th>
                                是否公開：
                            </th>
                            <td class="eTableContent">
                                <select name="fCTUPublic" id="fCTUPublic" runat="server">
                                    <option value="Y" selected="selected">公開</option>
                                    <option value="N">不公開</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                <font color="red">*</font> 學名：
                            </th>
                            <td class="eTableContent">
                                <input name="scientificName" id="scientificName" size="50" runat="server" value="(學名)" />
                            </td>
                        </tr>
                        <tr>
                            <th>
                                科中文名：
                            </th>
                            <td class="eTableContent">
                                <input name="family" id="family" size="50" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <th>
                                俗名：
                            </th>
                            <td class="eTableContent">
                                <input name="commonName" id="commonName" size="50" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <th>
                                <font color="red">*</font> 產期：
                            </th>
                            <td class="eTableContent">
                                <textarea name="season" id="season" cols="50" rows="7" runat="server" value="(產期)"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                <font color="red">*</font> 產地：
                            </th>
                            <td class="eTableContent">
                                <textarea name="origin" id="origin" cols="50" rows="7" runat="server" value="(產地)"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                台灣分佈：
                            </th>
                            <td class="eTableContent">
                                <textarea name="distributionInTaiwan" id="distributionInTaiwan" cols="50" rows="7"
                                    runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                棲息環境：
                            </th>
                            <td class="eTableContent">
                                <textarea name="habitats" id="habitats" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                參考文獻：
                            </th>
                            <td class="eTableContent">
                                <textarea name="reference" id="reference" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                形態特徵：
                            </th>
                            <td class="eTableContent">
                                <textarea name="characteristic" id="characteristic" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                棲所生態：
                            </th>
                            <td class="eTableContent">
                                <textarea name="habitatsType" id="habitatsType" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                地理分布：
                            </th>
                            <td class="eTableContent">
                                <textarea name="distribution" id="distribution" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                        <tr>
                            <th>
                                漁業利用：
                            </th>
                            <td class="eTableContent">
                                <textarea name="utility" id="utility" cols="50" rows="7" runat="server"></textarea>
                            </td>
                        </tr>
                    </table>
                    <table border="0" width="100%" cellspacing="0" cellpadding="0">
                        <tr>
                            <td align="center">
                                <input type="hidden" name="action_equals_Del" id="action_equals_Del" value="0" />
                                <input type="button" value="編修存檔" class="cbutton" onclick="edit()" />
                                <% if (id != 0)
                                   { %>
                                <input type="button" value="刪　除" class="cbutton" onclick="del()" />
                                <% } %>
                                <input type="reset" value="重　填" class="cbutton" />
                                <input type="button" value="回前頁" class="cbutton" onclick="location.href='<%=returnUrl%>'" />
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
        if (s == "0")
            window.location.href = "EditCrop.aspx";
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

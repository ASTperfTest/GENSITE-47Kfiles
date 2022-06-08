<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DsdASPXQuery.aspx.cs" Inherits="DsdASPXQuery" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>資料上稿-入口網</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<link href="../inc/setstyle.css" rel="stylesheet" type="text/css"/>
</head>
<body>
<form id="Form1" runat="server">
    <div>
        <div id="FuncName">
		    <table width="100%" border="0" cellSpacing="0" cellPadding="0">
			    <tr><td width="100%" class="FormName" colSpan="2">
			        資料管理／資料上稿-入口網<font size="2">【目錄樹節點: <%=Session["catName"]%>】【查詢】</font>
			    </td></tr>
			</table>
			<hr SIZE="1" color="#000080" noShade="noshade"/>
            <div id="Nav" align="right">
                <a style="font-size:9pt;"  href="DsdASPXList.aspx?ItemID=<%=Session["itemID"] %>&CtNodeID=<%=Session["ctNodeId"] %>" title="回前頁">回前頁</a>
            </div>
            <div id="ClearFloat">
            </div>
        </div>
        <div id="FormName" align="center">
            單元資料維護【主題單元:<%=Session["CtUnitName"]%> / 單元資料:純網頁】
        </div>
        <div>
		<table Width="90%" border="0" id="ListTable" cellspacing="1" cellpadding="2" align="center" class="eTable">
            <tr>
                <td align="center" width="15%" class="eTableLable">標題</td>
                <td class="eTableContent">
                    <asp:TextBox runat="server" ID="titleTxt" size="50"></asp:TextBox>
                </td>
            </tr>
		</table>
        </div>
         <table align="center">
            <tr>
                <td width="100%">
                    <asp:Button runat="server" ID="btnQuery" class="cbutton" Text="查    詢" 
                        onclick="btnQuery_Click" />
                    <input type="reset" id="btnRest" class="cbutton" value="重    填" onclick="returnToQuery()" />
                </td>
            </tr>
        </table>
    </div>
   
    </form>
    </body>
</html>

<script type="text/javascript">
        
        function returnToQuery() {
            window.location.href = "DsdASPXQuery.aspx";

        }
	
</script>

<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CatTreePrintContent.aspx.vb" Inherits="CatTreePrintContent" %>

<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/tr/html4/strict.dtd'>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>    
    <title>農業知識入口網 －小知識串成的大力量－/</title>
    <link rel="stylesheet" type="text/css" href="/xslgip/style3/css/4seasons.css" />
    <link rel="stylesheet" type="text/css" href="/xslgip/style3/css5/AM.css" />
</head>
<body onload="print();close();">
<br/>
<center>
    <table width="600" border="0">
        <tr>
            <td  align="left">
                <asp:Label runat="server" ID="TabText" Text=""></asp:Label>
                <div class="path">目前位置：&gt;
                    <asp:Label runat="server" ID="NavUrlText" Text=""></asp:Label></div>
                <h3>
                    <asp:Label runat="server" ID="NavTitleText" Text=""></asp:Label></h3>
                <div id="Magazine">
                    <div class="Event">
                        <asp:Label ID="TableText" runat="server" Text="" />
                    </div>
                </div>
                <div class="">
                    <%=relationsArticle%>
                </div>
            </td>
            <td><%=innerTag%></td>
        </tr>
    </table>
</center>    
<br/>
</body>
</html>

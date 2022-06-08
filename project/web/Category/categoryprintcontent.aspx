<%@ Page Language="C#" AutoEventWireup="true" CodeFile="categoryprintcontent.aspx.cs"
    Inherits="Category_categoryprintcontent" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>農業知識入口網 －小知識串成的大力量－/</title>
    <link rel="stylesheet" type="text/css" href="/xslgip/style3/css/4seasons.css" />
    <link rel="stylesheet" type="text/css" href="/xslgip/style3/css5/AM.css" />
</head>
<body onload="print();close();">
    <br />
    <center>
        <table width="600" border="0">
            <tr>
                <td align="left">
                    <asp:Label runat="server" ID="TabText" Text=""></asp:Label>
                    <div class="path">
                        目前位置：&gt;
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
            </tr>
        </table>
    </center>
    <br />
</body>
</html>

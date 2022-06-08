<%@ Page Language="VB" AutoEventWireup="false" CodeFile="knowledge_cpPrint.aspx.vb" Inherits="knowledge_cpPrint" Title="農業知識入口網-農業知識家" ValidateRequest="false" %>
<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>
<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/tr/html4/strict.dtd'>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>農業知識入口網 －小知識串成的大力量－/</title>
    <link rel="stylesheet" type="text/css" href="/xslgip/style3/css/4seasons.css" />
    <link rel="stylesheet" type="text/css" href="/xslgip/style3/css5/AM.css" />
</head>
<body onload="print();close();">
<form runat="server">
    <br />    
        <table width="600" border="0" align="center">
            <tr>
                <!--中間內容區-->
                <td class="center">
                    <uc2:tabtext id="TabText1" runat="server" />
                    <div class="path">
                        目前位置：<a href="/knowledge/knowledge.aspx">首頁</a><asp:Label ID="PathText" runat="server" Text=""></asp:Label></div>
                    <div class="depdate">
                        問題建置日期：<asp:Label ID="xPostDateText" runat="server" Text=""></asp:Label></div>
                    <!-- 頁面功能 Start -->
                    
                    <!-- 頁面功能 end -->
                    <!-- 內容區 Start -->
                    <div id="cp">
                        <!-- 內容區 Start -->
                        <asp:Label ID="QuestionText" runat="server" Text=""></asp:Label>
                        <!--駐站專家補充 /開始 -->
                        <asp:Label ID="ProsText" runat="server" Text=""></asp:Label>
                        <!--討論/開始 -->
                        <asp:Label ID="DiscussText" runat="server" Text=""></asp:Label>
                        <!--相似問題/開始 -->
                        <asp:Label ID="SimilarProblemText" runat="server" Text=""></asp:Label>
                    </div>
                    <!--20080829 chiung -->
                </td>
                <td align="left"><%'=innerTag%></td>
                <td align="left"><%'=accountTag%></td>
            </tr>
        </table>        
    </form>
</body>
</html>

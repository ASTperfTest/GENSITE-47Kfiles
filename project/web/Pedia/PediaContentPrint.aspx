<%@ Page Language="VB" AutoEventWireup="false" CodeFile="PediaContent.aspx.vb" Inherits="Pedia_PediaContent" Title="農業知識小百科" %>

<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01//EN' 'http://www.w3.org/tr/html4/strict.dtd'>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>農業知識入口網 －小知識串成的大力量－/</title>
    <link rel="stylesheet" type="text/css" href="/xslgip/style3/css/4seasons.css" />
    <link rel="stylesheet" type="text/css" href="/xslgip/style3/css5/AM.css" />
</head>
<body onload="print();close();">
    <link rel="stylesheet" href="../css/pedia.css" type="text/css" />
    <table width="600" border="0" align="center">
        <tr>
            <td align="left">
                <div class="path">
                    目前位置：<a href="/">首頁</a>&gt;<a href="/Pedia/PediaList.aspx">農業知識小百科</a></div>                
                <div class="pantology">
                    <div class="head">
                    </div>
                    <div class="body">
                        <h2>
                            農業知識小百科</h2>
                        <table class="type01" summary="結果條列式">
                            <tr>
                                <th width="15%">詞目</th>
                                <td>
                                    <asp:Label ID="WordTitle" runat="server" /></td>
                            </tr>
                            <tr style='display: none;'>
                                <th>英文詞目</th>
                                <td>
                                    <asp:Label ID="WordEngTitle" runat="server" /></td>
                            </tr>
                            <tr style='display: none;'>
                                <th>學名(中/英)</th>
                                <td>
                                    <asp:Label ID="WordFormalName" runat="server" /></td>
                            </tr>
                            <tr style='display: none;'>
                                <th>俗名(中/英)</th>
                                <td>
                                    <asp:Label ID="WordLocalName" runat="server" /></td>
                            </tr>
                            <tr style='display: none;'>
                                <th>名詞釋義</th>
                                <td>
                                    <asp:Label ID="WordBody" runat="server" /></td>
                            </tr>
                            <tr style='display: none;'>
                                <th>相關詞</th>
                                <td>
                                    <asp:Label ID="WordKeyword" runat="server" /></td>
                            </tr>
                        </table>
                        <hr />
                        <ul class="Function2" runat="server" visible="false">
                            <li>
                                <asp:HyperLink ID="CanAddLink" runat="server" CssClass="add" Visible="false">我來補充</asp:HyperLink></li>
                            <li>
                                <asp:HyperLink ID="ExplainListLink" runat="server" CssClass="add_more" Visible="false">看全部</asp:HyperLink></li>
                        </ul>
                        <h2>
                            補充解釋：</h2>
                        <asp:Label ID="ExplainText" runat="server" />
                        <div class="top">
                            <a href="#" title="top">top</a></div>
                    </div>
                    <div class="foot">
                    </div>
                </div>
                <!-- 內容區 End -->
                <%=accountTag%>
        </tr>
    </table>

    <script language="javascript" type="text/javascript">
        function WordSearch(word) {
            document.SearchForm.Keyword.value = word;
            document.SearchForm.action = "/kp.asp?xdURL=Search/SearchResultList.asp&mp=1";
            document.SearchForm.submit();
        }
    </script>

</body>
</html>

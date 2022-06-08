<%@ Page Language="C#" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true"
    CodeFile="Recommand_Add.aspx.cs" Inherits="Recommand_Add" Title="農業知識入口網-好文推薦" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <link rel="stylesheet" href="../js/multiselect/css/common.css" type="text/css" />
    <link type="text/css" href="../js/multiselect/css/ui.multiselect.css" rel="stylesheet" />
    <link type="text/css" rel="stylesheet" href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.8.10/themes/ui-lightness/jquery-ui.css" />
    <script type="text/javascript" src="../js/multiselect/js/jquery-1.6.1.min.js"></script>
    <script type="text/javascript" src="../js/multiselect/js/jquery-ui-1.8.14.custom.min.js"></script>
    <script type="text/javascript" src="../js/multiselect/js/ui.multiselect.js"></script>
    <script language="javascript" type="text/javascript">
        $(function () {
            $(".multiselect").multiselect();
        });
    </script>
    <!--中間內容區-->
    <div class="path" style="float: left; padding-top: 10px">
        目前位置：
    </div>
    <div style="width: 80%; float: left;">
        <ul id="path_menu">
            <li><a href="../mp.asp?mp=1">首頁</a></li>
            <li style='top: 10px;'>></li>
            <li><a href="Recommand_Add.aspx">好文推薦</a></li>
        </ul>
    </div>
    <div style="float: left; width: 100%;">
        <table cellpadding="2" cellspacing="4" style="width:100%; font-size:1.25em">
            <tr>
                <td colspan="2">
                    <h3>好文推薦</h3>
                </td>
            </tr>
            <tr>
                <td colspan="2" style=" line-height:20px">
                    此功能提供會員推薦相關農業新知等文章，<b>經系統管理者確認及審查通過後</b>將公告於本網站上，<b>管理者並保留修改其內容之權利</b>。為獎勵會員，推薦網站通過審查後，此行為亦納入會員KPI計分。<br/>
                    <p><br/><b>推薦文章公告說明：</b><br/>
                    <ol>
                        <li>推薦文章必須為『農業知識』類別文章，如：農業故事、新知介紹、栽培技術、防治方法、病蟲害防制技術...等；『新聞』、『廣告』、『政治』等相關文章則不符合本單元範圍。此外，您推薦的文章必須是該文章所有權人同意於網路上公開點閱的文章，禁止推薦未經授權的文章資料。</p></li>
                        <li>在各位會員進行推薦前，先利用『查詢文章』功能，搜尋文章是否已被推薦。重複推薦的資料將會被網站管理人員刪除。</li>
                        <li>若大量推薦不符合本單元範圍的文章、重覆文章、灌水文章等，或被檢舉文章來源不合法、或刻意以不正當的行為獲取積分...等影響本功能正常營運之行為，情節嚴重者，管理者保留扣除會員得點和停止會員資格的權力。</li>
                    </ol>
                    
                </td>
            </tr>
            <tr>
                <td colspan="2" style="line-height: 20px" align="center">
                    <table  style="background-color: #ffc; width:90%;" cellpadding="2" cellspacing="3">
                        <tr>
                            <td align="left">
                                <font color="red">*</font><b>文章標題</b>：<br />
                                <asp:textbox id="txtTitle" runat="server" MaxLength="50" autocompletetype="Disabled"  width="90%"></asp:textbox>
                            </td>
                        </tr>
                        <tr>
                            <td align="left">
                                <font color="red">*</font><b>推薦好文網址</b>：<br />
                                <asp:textbox runat="server" id="txtUrl" MaxLength="200" autocompletetype="Disabled"  width="90%"></asp:textbox>
                            </td>
                        </tr>
                        <tr>
                            <td align="left">
                                <font color="red">*</font><b>資料出處</b>(請輸入資料來源資訊，例如：wikipedia、xxx部落客 等等)：<br />
                                <asp:textbox id="txtSource" runat="server" width="90%" MaxLength="50" autocompletetype="Disabled" ></asp:textbox>
                            </td>
                        </tr>

                        <tr>
                            <td align="left">
                                <font color="red">*</font><b>推薦原因</b>(必須介於20至300個字元)：<br />
                                <asp:textbox id="txtContent" runat="server" width="90%" textmode="MultiLine" rows="5"></asp:textbox>
                            </td>
                        </tr>                        
                        <tr>
                            <td align="left">
                                <font color="red">*</font><b>分類</b>：<br /> 
                                <asp:CheckBoxList ID = "listTAGs" runat = "server" RepeatDirection="Horizontal"  ></asp:CheckBoxList> 
                                
                                <%--<asp:Checkbox runat="server" id="listTAGs" cssclass="multiselect" style="width:50%;"
                                    selectionmode="Multiple" on></asp:Checkbox>--%>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" align="center">
                                <asp:button id="btnConfirm" runat="server" text="確定推薦" onclick="btnConfirm_Click" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>                
        </table>
    </div>    
    
    <script type="text/javascript">
        var errString = '<%=errString %>'
        var saveOK = '<%=saveOK %>'
        if (errString != '') {
            alert(errString);
        }
        if (saveOK == "Y") {
            alert('謝謝您的推薦，審核須要一些時間，審核通過後即會顯示於[好文推薦]功能中。');
            window.location.href = '/recommand/recommand_Mylist.aspx?T=<%=DateTime.Now.Ticks %>'
        }
    </script>
</asp:Content>

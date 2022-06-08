<%@ Page Title="" Language="C#" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true"
    CodeFile="MemberInvitePage.aspx.cs" Inherits="MemberInvitePage" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <%--複製 textarea.value--%>
    <script type="text/javascript">
        function checkMailAddress() {
            var txt1 = document.getElementById('<%=txtAddress_1.ClientID %>').value.replace(/^[\s　]+/g, "").replace(/[\s　]+$/g, "");
            var txt2 = document.getElementById('<%=txtAddress_2.ClientID %>').value.replace(/^[\s　]+/g, "").replace(/[\s　]+$/g, "");
            var txt3 = document.getElementById('<%=txtAddress_3.ClientID %>').value.replace(/^[\s　]+/g, "").replace(/[\s　]+$/g, "");
            var txt4 = document.getElementById('<%=txtAddress_4.ClientID %>').value.replace(/^[\s　]+/g, "").replace(/[\s　]+$/g, "");
            if (txt1 == "" && txt2 == "" && txt3 == "" && txt4 == "") {
                alert("至少輸入一位好友的信箱");
                document.getElementById('<%=txtAddress_1.ClientID %>').value = "";
                document.getElementById('<%=txtAddress_2.ClientID %>').value = "";
                document.getElementById('<%=txtAddress_3.ClientID %>').value = "";
                document.getElementById('<%=txtAddress_4.ClientID %>').value = "";
                event.returnValue = false;
            }

        }
        function copyToClipboard(txt) {
            var copied = false;
            if (window.clipboardData) {
                window.clipboardData.clearData();
                window.clipboardData.setData("Text", txt);
                copied = true;
            } else if (navigator.userAgent.indexOf("Opera") != -1) {
                window.location = txt;
            } else if (window.netscape) {
                try {
                    netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");
                } catch (e) {
                    alert("被瀏覽器拒絕！\n請在瀏覽器地址欄輸入'about:config'並回上頁\n然後將'signed.applets.codebase_principal_support'設置為'true'");
                }
                var clip = Components.classes['@mozilla.org/widget/clipboard;1'].createInstance(Components.interfaces.nsIClipboard);
                if (!clip)
                    return;
                var trans = Components.classes['@mozilla.org/widget/transferable;1'].createInstance(Components.interfaces.nsITransferable);
                if (!trans)
                    return;
                trans.addDataFlavor('text/unicode');
                var str = new Object();
                var len = new Object();
                var str = Components.classes["@mozilla.org/supports-string;1"].createInstance(Components.interfaces.nsISupportsString);
                var copytext = txt;
                str.data = copytext;
                trans.setTransferData("text/unicode", str, copytext.length * 2);
                var clipid = Components.interfaces.nsIClipboard;
                if (!clip)
                    return false;
                clip.setData(trans, null, clipid.kGlobalClipboard);
                copied = true;
            }
            if (copied) alert('已經複製');
            else alert("使用的瀏覽器不支援文字複製功能!");
        }
    </script>
    <style type="text/css">
        #divWrapper
        {
            
        }
        #divEmail
        {
            float: left;
            width: 360px;
        }
        #divEmail table
        {
            width: 80%;
            text-align: center;
        }
        #divEmail table [type=text]
        {
            width: 250px;
            margin-bottom:2px;
        }
        #divFB
        {
            float: left;
        }
        #divURL
        {
            clear: both;
            text-align: center;            
        }
        textarea
        {
            overflow: hidden;
        }
        
        #InviteeTable {   
          border: 1px solid #c6c6c6;   
          border-collapse: collapse;   
        }   
        #InviteeTable tr,#InviteeTable td {   
          border: 1px solid #c6c6c6;   
          height:20px;
          color:black;
        }
        #InviteeTable #header {   
           font-weight:bold;
           color:black;
        }        
    </style>    
    <div class="path" style="float: left; padding-top: 10px">目前位置：</div>
    <div style="float: left; width: 80%">
        <ul id="path_menu">
            <li><a title="首頁" href="/mp.asp?mp=1">首頁 </a></li>
            <li style="top: 10px;">&gt;</li>
            <li><a href="/Member/MemberInvitePage.aspx">邀請好友</a></li>
        </ul>
    </div>
    <h3>邀請好友-相招來農業知識入口網 腦袋充電拿3C</h3>
    <div id="divWrapper" style="margin:0px 10x 5px 10px">
        <div id="divHeader">            
        </div>
        <div id="divBody">
            <p style=" margin:5px 10px; line-height:20px;">
                你知道台灣路邊隨處可見的芒草，因具有發電減碳之功能被歐洲科學家視為珍寶嗎？<br/>
                你知道南台灣盛產的芒果，在日本、韓國是上等高級品嗎？</p>
            <p style=" margin:10px 10px; line-height:20px;">
                馬上邀請好友一同加入農業知識入口網會員吧！成為正港的台灣神農氏，相招加入會員還可衝點數、玩活動抱3C大獎回家！
                還有還有，你的朋友只要成功加入會員，我們還額外送你40點KPI分數喔！！<br/>
                還在等什麼呢？快點動動手邀請好友加入會員吧～
            </p><br/>
            <b>邀請好友公告說明</b>
            <ol style=" margin:10px 10px 20px;; line-height:20px;">
                <li>推薦的朋友必須成功加入會員並完成Email認證程序後才能獲得KPI點數；當推薦的朋友被停權，先前獲得的KPI點數將會被扣回。</li>
                <li>若刻意以不正當的行為獲取積分(如：大量註冊灌水會員)或以其他方式影響本功能正常營運之行為，情節嚴重者，管理者保留扣除會員得點和停止會員資格的權力。</li>
            </ol>
        </div>
        <div style=" background-color:#FFC;">
            <div id="divEmail">
                <table>
                    <tr align="left">
                        <td><b>邀請方式1：</b><br/>寄送邀請信，輸入好友的 E-Mail address:</td>
                    </tr>
                    <tr>
                        <td>
                            <asp:Panel runat="server" ID="panelMail">
                                Email 1:<asp:TextBox runat="server" ID="txtAddress_1" Text=""></asp:TextBox><br/>
                                Email 2:<asp:TextBox runat="server" ID="txtAddress_2" Text=""></asp:TextBox><br/>
                                Email 3:<asp:TextBox runat="server" ID="txtAddress_3" Text=""></asp:TextBox><br/>
                                Email 4:<asp:TextBox runat="server" ID="txtAddress_4" Text=""></asp:TextBox><br/>
                                    
                                    <asp:Literal runat="server" ID="litContent"></asp:Literal>                                 
                            </asp:Panel>
                            
                            <asp:Button runat="server" ID="btnSubmit" Text="送出" OnClientClick="checkMailAddress()"
                                OnClick="btnSubmit_Click" />                            
                        </td>
                    </tr>
                </table>
            </div>
            <div id="divFB">
                <table>
                    <tr>
                        <td><b>邀請方式2：</b><br/>發佈訊息到Facebook:</td>
                    </tr>
                    <tr>
                        <td><img alt="農業知識入口網" title="農業知識入口網" src="/images/FBIcon.png" /></td>
                        <td><textarea id="txtFBContent" rows="6" cols="30"  style="resize:none">朋友們，介紹你們一個好站，可以讓你輕鬆習得農業知識，掌握最新農業資訊，快來加入農業知識入口網吧～～</textarea></td>
                    </tr>
                    <tr>
                        <td></td>
                        <td><input type="button" value="發佈" onclick="<%=string.Format("window.open('{0}/Member/{1}?v=2&Scode={2}');return false;", strURL, "MIP.aspx", InvitationCode) %>" /></td>
                    </tr>
                </table>
            </div>
            <div id="divURL" style="text-align:left">
                <br />
                <p><b>邀請方式3：</b><br/>
                   複製此網址寄給您的朋友們：<asp:TextBox runat="server" ID="txtInvitationCode" Width="450px" ReadOnly="true"></asp:TextBox>
                    <input type="button" id="btnCopy" value="複製" onclick="copyToClipboard(<%=txtInvitationCode.ClientID %>.value)" />
                </p><br/>
            </div>
        </div>        
        <br/>
        <div style="background-color:#e3f1fe; padding:20px 0px 20px 0px;">
        <b>已邀請成功的朋友清單：</b><br/>
        <table style="margin:5px 0px 0px 20px ; width:450px; border-color:#c6c6c6" id="InviteeTable" border="1" cellpadding="1" cellspacing="0">
            <tr id="header">
                <td>&nbsp;&nbsp;</td>
                <td>朋友暱稱</td>
                <td>註冊時間</td>
                <td>是否通過認證</td>
            </tr>
            <asp:Label runat="server" ID="labInvitation" Text=""></asp:Label>
        </table>
        </div>
    </div>
    
    <script type='text/javascript'>
        var sendOK = '<%=sendOK %>';
        if (sendOK == 'Y') {
            alert('信送出囉');
            location.href = '/Member/MemberInvitePage.aspx';
        }
    </script>
</asp:Content>

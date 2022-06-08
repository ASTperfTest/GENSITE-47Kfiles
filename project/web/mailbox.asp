<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <title>農業知識入口網 －小知識串成的大力量－</title>
    <base target="_self" />

    <script language="javascript" type="text/javascript">
        var title = "";
        function SendMail() {

            var url = window.dialogArguments.document.URL;
            var verify = document.getElementById('CheckCodeforMail').value;
            var objurl = document.getElementById('urlRecord');
            objurl.value = url;
            var typeThisPage = document.getElementById('type');
            var articleThisPage = document.getElementById('ARTICLE_ID');
            var getURL = document.URL;            
            var arr = getURL.split("?");
            
            if (arr.length > 1) {
                if (arr[1].indexOf("Discuss", 0) >= 0) {
                    title = "檢舉信";                                        
                    typeThisPage.value = "2";
                    var article = window.dialogArguments.document.getElementById('ARTICLE_ID');

                    var DiscussID = <%=request("DiscussID") %>

                    document.getElementById('account').value = window.dialogArguments.document.getElementById('d_account_' + DiscussID).value;
                    document.getElementById('nickname').value = window.dialogArguments.document.getElementById('d_nickname_' + DiscussID).value;
                    
                    articleThisPage.value = article.value;
                }
                else if (arr[1].indexOf("Question", 0) >= 0) {
                    title = "系統問題反應";
                    typeThisPage.value = "1";
                    var articleQuestion = window.dialogArguments.document.getElementById('ctl00_ContentPlaceHolder1_ARTICLE_ID_Question');
                    articleThisPage.value = articleQuestion.value;
                }
                else {
                    title = "系統問題反應";
                    
                    var type = window.dialogArguments.document.getElementById('type');
                    typeThisPage.value = type.value;
                    var article = window.dialogArguments.document.getElementById('ARTICLE_ID');
                    articleThisPage.value = article.value;
                }
            }
            else {
                title = "系統問題反應";
                
                var type = window.dialogArguments.document.getElementById('type');
                typeThisPage.value = type.value;
                var article = window.dialogArguments.document.getElementById('ARTICLE_ID');
                articleThisPage.value = article.value;
            }           
            
            document.getElementById('h2Title').innerHTML = title;
        }
    </script>



    <link href="/xslGip/style3/css/4seasons.css" rel="stylesheet" type="text/css" />
</head>
<body onload="SendMail()">
    <div class="pantology">
        <div class="head"></div>
        <div class="body">
            <h2 id="h2Title"></h2>
            <form name="form1" method="post" id="form1" action="SendMail.asp">
                <table class="type01">
                    <tr>
                        <th nowrap="nowrap" scope="row" style="font-size: 0.75em; width: 100px;">
                            發表您的意見:
                        </th>
                        <td style='padding-right:30px;'>
                            <textarea name="txtDisCussion" rows="10"  id="txtDisCussion"><%=session("SendMail_txtDisCussion") %></textarea>
                        </td>
                    </tr>
                    <tr>
                        <th rowspan="2" nowrap="nowrap" scope="row" style="font-size: 0.75em; width: 100px;">
                            機器人辨識:
                        </th>
                        <td>
                            <image name="checkcode" src="/VerifyImageforMail.asp" width="80pt" style="border-bottom: 1px solid #999999;
                                border-top: 1px solid #999999; border-left: 1px solid #999999; border-right: 1px solid #999999;" onClick="this.src='VerifyImageforMail.asp?dumy=' + Math.random()" />
							<img name="reloadimg" alt="更新驗證碼" src="images/CheckCodeReload/reload.jpg" onclick="document.all.checkcode.src='VerifyImageforMail.asp?dumy=' + Math.random();" />
						</td>

                    </tr>
                    <tr>
                        <td>
                            <input type="text" name="CheckCodeforMail" id="CheckCodeforMail" /><br />
                            請輸入上方圖片的數字
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <input type="hidden" name="account" id="account" /><br />
                            <input type="hidden" name="nickname" id="nickname" />
                        </td>
                    </tr>
                </table>
                <div class="s01">
                    <input type="submit" name="ButtonSave" value="確認" id="ButtonSave" />
                    <input type="button" name="ButtonCancle" value="取消" id="ButtonCancle" onclick="if(!window.confirm('確定要離開嗎？')){return false;} window.close();" />
                </div>
                <input type="hidden" name="urlRecord" id="urlRecord" />
                <input type="hidden" name="type" id="type" />
                <input type="hidden" name="ARTICLE_ID" id="ARTICLE_ID" />
            </form>
        </div>
        <div class="foot">
        </div>
    </div>
</body>
</html>

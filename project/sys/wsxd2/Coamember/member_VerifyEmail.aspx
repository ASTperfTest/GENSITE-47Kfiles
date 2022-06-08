
<script language="JavaScript">
<!--
function send() {
	var form = document.Form1; 
	if(form.email.value==""){
		alert("您忘了填寫 E-mail 了！"); 
		form.email.focus(); 
		return false;
	}
	else if(form.email.value.indexOf("@")<=-1){
		alert("您所填寫的 E-mail 格式有誤！"); 
		form.email.focus(); 
		return false;
	}
	else if(form.email.value.indexOf("'")>=0){
		alert("您所填寫的[ E-mail ]格式有誤！"); 
		form.email.focus(); 
		return false;
	}
	if(form.captcha.value==""){
		alert("您忘了填寫 驗證碼 了！"); 
		form.captcha.focus(); 
		return false;
	}
	form.email.value = escape(form.email.value);
	return true;
}  
//-->
</script>

<div>
    <div class="path">
        目前位置：<a title="首頁" href="mp.asp">首頁</a>&gt; 加入會員</div>
    <h3>
        <b>農業知識入口網會員註冊系統</b></h3>
    <div id="Magazine">
        <div class="Event">
            <div class="experts">
                <div>
                    註冊開始，先讓我們確認您的 E-mail 信箱有效。
                    <br />
                    <ol>
                        <li>Step.1 確認信箱。</li>
                        <li>Step.2 確定送出後，系統會發送一封 E-mail 的認證信到您所填的 E-mail 帳號中。</li>
                        <li>Step.3 請您到這個 E-mail 信箱中收取，並確認無誤後，按 [E-mail 確認無誤] 回覆。</li>
                        <li>Step.4 填寫註冊資料。</li>
                        <li>Step.5 完成！</li>
                    </ol>
                    <br />
                </div>
                <form name="Form1" id="Form1" method="post" class="FormA" action="sp.asp?xdURL=coamember/member_VerifyEmailPost.aspx" onsubmit="return send()">
				<input id="returnUrl" name="returnUrl" type="hidden" value="<%= Request("returnUrl")%>" />
				<table cellspacing="0" class="DataTb1" width="100%">
                    <tr>
                        <th align="right" height="30">
                            <label for="email">
                                *E-mail&nbsp;</label>
                        </th>
                        <td colspan="2">
                            <input name="email" type="text" value="" class="Text" id="email" size="30" maxlength="50" />
                        </td>
                    </tr>
                    <tr>
                        <th align="right" height="30">
                            <label for="captcha">
                                *圖形驗證碼&nbsp;</label>
                        </th>
                        <td colspan="2">
                            <input name="captcha" type="text" value="" class="Text" id="captcha" size="30" />
                            <image src="/VerifyImageforMail.asp" align="absmiddle" height="25" width="80pt" style="border-bottom: 1px solid #999999;
                                border-top: 1px solid #999999; border-left: 1px solid #999999; border-right: 1px solid #999999;" />
                            &nbsp;(請輸入圖片中的字母)
                        </td>
                    </tr>
                    <tr>
                        <td align="center" colspan="2">
                            <br />
                            <input id="Submit" name="Submit" type="submit" class="Button" value="確定" />
                            <input id="Cancel" name="Cancel" type="button" class="Button" value="取消" />
                        </td>
                    </tr>
                </table>
                </form>
            </div>
        </div>
    </div>
</div>

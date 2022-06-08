<!--#Include virtual = "/inc/client.inc" -->
<%
	Dim SMemberId, RMemberId, updateMemID, Flag
	
	SMemberId = Session("memID")
	RMemberId = Request("memID")
	updateMemID = Request("updateMemID")
	Flag = true
	
	If SMemberId = "" Or RMemberId = "" Then
		Flag = false
	End If
	
	If SMemberId <> RMemberId Then
		Flag = false
	End If
	
	If updateMemID<>"" Then		
		Flag = true
		SMemberId = updateMemID		
	End If
	
	If Flag = false Then		
		response.write "<script> alert('尚未登入！');window.location.href='mp.asp';</script>"
	End If

	Dim account, email
	
	if Flag = true then
		Set rsreg = conn.Execute("select account,email from Member Where account = '" & SMemberId & "' ")	
		
		if not rsreg.eof then
			account = trim(rsreg("account"))
			email = trim(rsreg("email"))
		end if
	end if
%>
<script language="JavaScript">
<!--
function send() {
	var form = document.Form1; 
	if(form.email.value==""){
		alert("您忘了填寫[ E-mail ]了！"); 
		form.email.focus(); 
		return false;
	}
	else if(form.email.value.indexOf("@")<=-1){
		alert("您所填寫的[ E-mail ]格式有誤！"); 
		form.email.focus(); 
		return false;
	}
	else if(form.email.value.indexOf("'")>=0){
		alert("您所填寫的[ E-mail ]格式有誤！"); 
		form.email.focus(); 
		return false;
	}
	if(form.password.value==""){
		alert("您忘了輸入[ 密碼 ]了！"); 
		form.password.focus(); 
		return false;
	}
	form.email.value = escape(form.email.value);
	return true;
}  
//-->
</script>

<div>
    <div class="path">
        目前位置：<a title="首頁" href="mp.asp">首頁</a>&gt; 重新認證 E-mail</div>
    <h3>
        <b>會員 E-mail 認證</b></h3>
    <div id="Magazine">
        <div class="Event">
            <div class="experts">
                <div>
                    提醒您至加入會員時所填的 Email 信箱收認證信。</br>
					若您沒有收到認證信，為了確保您的 E-mail 是有效以保障您的權利，請按以下步驟重新確認您的 E-mail 後才可以登入農業知識入口網
                    <br />
                    <ol>
                        <li>請重新輸入您的 E-mail ，及登入密碼。</li>
                        <li>確定送出後，系統會發送一封 E-mail 的認證信到您所填的 E-mail 帳號中。</li>
                        <li>請您到這個 E-mail 信箱中收取，並確認無誤後，按 [E-mail 確認無誤] 回覆。</li>
                    </ol>
                    <br />
					
                </div>
                <form name="Form1" id="Form1" method="post" class="FormA" action="sp.asp?xdURL=coamember/member_reVerifyEmailPost.aspx" onsubmit="return send()">
                <table cellspacing="0" class="DataTb1" width="100%">
                    <tr>
                        <th align="right" height="30">
                            <label for="account">農業知識入口網帳號&nbsp;</label>
                        </th>
                        <td class="modifytd" colspan="2">
						<input id="account" name="account" type="hidden" value="<%= Server.HTMLEncode(account)%>" />
						<%= Server.HTMLEncode(account)%></td>

                    </tr>
                    <tr>
                        <th align="right" height="30">
                            <label for="password">密碼&nbsp;</label>
                        </th>
                        <td colspan="2">
                            <input name="password" type="password" value="" class="Text" id="password" size="30" maxlength="30" />&nbsp;<br/>
                            (請再輸入一次密碼用以驗證您的身份)
                        </td>
                    </tr>
                    <tr>
                        <th align="right" height="30">
                            <label for="email">
                                *E-mail&nbsp;</label>
                        </th>
                        <td colspan="2">
                            <input name="email" type="text" value="<%= Server.HTMLEncode(email)%>" class="Text" id="email" size="30" maxlength="50" />&nbsp;<br/>
                            (系統會以此 E-mail 當作您最新的 E-mail 資料)
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

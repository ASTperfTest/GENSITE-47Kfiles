<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="newmember"
HTProgPrefix = "newMember"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/client.inc" -->
<!--#include file="newMemberMail.inc" -->
<!--#include virtual = "/inc/checkGIPAPconfig.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html>
<head>
    <style type="text/css">
        #InviteeTable
        {
            border: 1px solid #c6c6c6;
            border-collapse: collapse;
        }
        #InviteeTable tr, #InviteeTable td
        {
            border: 1px solid #c6c6c6;
            height: 20px;
            color: black;
        }
        #InviteeTable #header
        {
            font-weight: bold;
            color: black;</style>

<% 
	dim id
	id = trim(request.querystring("account"))
	
	gardenExpert_sql = "SELECT * FROM GARDENING_EXPERT WHERE  ACCOUNT= "& "'"  & id & "'"
	Set gardenExpert_RSreg = conn.execute(gardenExpert_sql)
	if not gardenExpert_RSreg.eof then
		gardenExpert_intro = gardenExpert_RSreg("INTRODUCTION")
		gardenExpert_order = gardenExpert_RSreg("SORT_ORDER")
	else
		gardenExpert_intro = ""
		gardenExpert_order = ""
	end if
	
	
	


	if request.querystring("Status")="Suspension" then	  
        sql = ""
        sql = sql & vbcrlf & "declare @accountId varchar(50)"
        sql = sql & vbcrlf & "set @accountId = '" & id & "'"
        sql = sql & vbcrlf & ""
        sql = sql & vbcrlf & "declare @inviterId varchar(50)"
        sql = sql & vbcrlf & "declare @KPIAddedSNO int"
        sql = sql & vbcrlf & "declare @Inviter_KPIPointe int"
        sql = sql & vbcrlf & ""
        sql = sql & vbcrlf & "if exists (select * from member where account=@accountId and [status]='Y')"
        sql = sql & vbcrlf & "begin"
        sql = sql & vbcrlf & "	begin tran"
        sql = sql & vbcrlf & "		select "
        sql = sql & vbcrlf & "			 @inviterId=InviteFriends_Head.Account"
        sql = sql & vbcrlf & "			,@KPIAddedSNO=InviteFriends_Detail.KPIAddedSNO"
        sql = sql & vbcrlf & "			,@Inviter_KPIPointe=Inviter_KPIPointe"
        sql = sql & vbcrlf & "		from InviteFriends_Detail "
        sql = sql & vbcrlf & "		inner join InviteFriends_Head on InviteFriends_Head.InvitationCode=InviteFriends_Detail.InvitationCode"
        sql = sql & vbcrlf & "		where inviteAccount = @accountId"
        sql = sql & vbcrlf & "		and IsActive = 1"
        sql = sql & vbcrlf & ""
        sql = sql & vbcrlf & "		--KPI扣分"
        sql = sql & vbcrlf & "		update MemberGradeLogin"
        sql = sql & vbcrlf & "			set loginInvite = (loginInvite-@Inviter_KPIPointe)"
        sql = sql & vbcrlf & "		from MemberGradeLogin "
        sql = sql & vbcrlf & "		where memberId=@inviterId "
        sql = sql & vbcrlf & "		and sno = @KPIAddedSNO"
        sql = sql & vbcrlf & "		and year(loginDate) = year(getdate()) --同年度的扣分，跨年度不予以追究(技正)"
        sql = sql & vbcrlf & ""
        sql = sql & vbcrlf & "		if (@@rowcount>0)"
        sql = sql & vbcrlf & "		begin			"
        sql = sql & vbcrlf & "			--標記停權(同年度才標記)"
        sql = sql & vbcrlf & "			update InviteFriends_Detail "
        sql = sql & vbcrlf & "				set IsActive = 9"
        sql = sql & vbcrlf & "			from InviteFriends_Detail "
        sql = sql & vbcrlf & "			where inviteAccount = @accountId and IsActive = 1			"
        sql = sql & vbcrlf & "		end"
        sql = sql & vbcrlf & "		"
        sql = sql & vbcrlf & "		UPDATE Member SET Status = 'N' WHERE account=@accountId"
        sql = sql & vbcrlf & "	commit"
        sql = sql & vbcrlf & "end"	  	  
		conn.execute(sql)
	end if 

	if request.querystring("Status")="RecoverRight" then
        sql = ""
        sql = sql & vbcrlf & "declare @accountId varchar(50)"
        sql = sql & vbcrlf & "set @accountId = '" & id & "'"
        sql = sql & vbcrlf & ""
        sql = sql & vbcrlf & "declare @inviterId varchar(50)"
        sql = sql & vbcrlf & "declare @KPIAddedSNO int"
        sql = sql & vbcrlf & "declare @Inviter_KPIPointe int"
        sql = sql & vbcrlf & ""
        sql = sql & vbcrlf & "if exists (select * from member where account=@accountId and [status]!='Y')"
        sql = sql & vbcrlf & "begin"
        sql = sql & vbcrlf & "	begin tran"
        sql = sql & vbcrlf & "		select "
        sql = sql & vbcrlf & "			 @inviterId=InviteFriends_Head.Account"
        sql = sql & vbcrlf & "			,@KPIAddedSNO=InviteFriends_Detail.KPIAddedSNO"
        sql = sql & vbcrlf & "			,@Inviter_KPIPointe=Inviter_KPIPointe"
        sql = sql & vbcrlf & "		from InviteFriends_Detail "
        sql = sql & vbcrlf & "		inner join InviteFriends_Head on InviteFriends_Head.InvitationCode=InviteFriends_Detail.InvitationCode"
        sql = sql & vbcrlf & "		where inviteAccount = @accountId"
        sql = sql & vbcrlf & "		and IsActive = 9"
        sql = sql & vbcrlf & ""
        sql = sql & vbcrlf & "		--KPI加分"
        sql = sql & vbcrlf & "		update MemberGradeLogin"
        sql = sql & vbcrlf & "			set loginInvite = (loginInvite+@Inviter_KPIPointe)"
        sql = sql & vbcrlf & "		from MemberGradeLogin "
        sql = sql & vbcrlf & "		where memberId=@inviterId "
        sql = sql & vbcrlf & "		and sno = @KPIAddedSNO		"
        sql = sql & vbcrlf & ""
        sql = sql & vbcrlf & "		if (@@rowcount>0)"
        sql = sql & vbcrlf & "		begin						"
        sql = sql & vbcrlf & "			update InviteFriends_Detail "
        sql = sql & vbcrlf & "				set IsActive = 1"
        sql = sql & vbcrlf & "			from InviteFriends_Detail "
        sql = sql & vbcrlf & "			where inviteAccount = @accountId and IsActive = 9"
        sql = sql & vbcrlf & "		end"
        sql = sql & vbcrlf & "		"
        sql = sql & vbcrlf & "		UPDATE Member SET Status = 'Y' WHERE account=@accountId"
        sql = sql & vbcrlf & "	commit"
        sql = sql & vbcrlf & "end"	  	   
		conn.execute(sql)
	end if 

	if request.querystring("scholar") = "pass" then
		sql = "UPDATE Member SET scholarValidate = 'Y', email = '" & request("email") & "' WHERE account = '" & id & "'"
		conn.execute(sql)
		Dim myname, myemail
		sql = "SELECT realname, email FROM Member WHERE account = '" & id & "'"
		set rs = conn.execute(sql)
		if not rs.eof then
			myname = rs("realname")
			myemail = rs("email")
		end if
		rs.close
		set rs = nothing
		if trim(myemail) <> "" then
			ePaperTitle = "農業知識入口網學者會員審核狀況通知"

			Body = "親愛的 " & myname & " 您好:<br/><br/>" & vbCrLf & vbCrLf
			Body = Body & "感謝您申請「農業知識入口網」的學者會員。<br/>" & vbCrLf
			Body = Body & "「農業知識入口網」""通過""您的學者會員申請。<br/>"& vbCrLf
			Body = Body & "有了您的參與將使得這個農業知識平台更理想更茁壯。我們熱切期盼您提供寶貴的意見與建議。<br/>" & vbCrLf
			Body = Body & "謝謝!<br/>" & vbCrLf
			Body = Body & "                                      敬祝平安<br/>" & vbCrLf
			Body = Body & "                                                系統管理員 敬上<br/>" & vbCrLf

			'S_Email = "taft_km@mail.coa.gov.tw"
			S_Email = """"&getGIPApconfigText("EmailFromName")&""" <"&getGIPApconfigText("EmailFrom")&">"
			R_Email = myemail
		
			Call Send_Email(S_Email,R_Email,ePaperTitle,Body)
		end if
	end if 
	
	if request.querystring("scholar") = "notpass" then
		sql = "UPDATE Member SET scholarValidate = 'N', email = '" & request("email") & "' WHERE account = '" & id & "'"
		conn.execute(sql)
		sql = "SELECT realname, email FROM Member WHERE account = '" & id & "'"
		set rs = conn.execute(sql)
		if not rs.eof then
			myname = rs("realname")
			myemail = rs("email")
		end if
		rs.close
		set rs = nothing
		if trim(myemail) <> "" then
			ePaperTitle = "農業知識入口網學者會員審核狀況通知"

			Body = "親愛的 " & myname & " 您好:<br/><br/>" & vbCrLf & vbCrLf
			Body = Body & "感謝您申請「農業知識入口網」的學者會員。<br/>" & vbCrLf
			Body = Body & "「農業知識入口網」""不通過""您的學者會員申請。<br/>"& vbCrLf
			Body = Body & "有了您的參與將使得這個農業知識平台更理想更茁壯。我們熱切期盼您提供寶貴的意見與建議。<br/>" & vbCrLf
			Body = Body & "謝謝!<br/>" & vbCrLf
			Body = Body & "                                      敬祝平安<br/>" & vbCrLf
			Body = Body & "                                                系統管理員 敬上<br/>" & vbCrLf

			'S_Email = "taft_km@mail.coa.gov.tw"
			S_Email = """"&getGIPApconfigText("EmailFromName")&""" <"&getGIPApconfigText("EmailFrom")&">"
			R_Email = myemail
		
			Call Send_Email(S_Email,R_Email,ePaperTitle,Body)
		end if
	end if 


	
	
	
	sql = "SELECT * FROM member WHERE  account= "& "'"  & id & "'"
	Set RSreg = conn.execute(sql)
	
	'檢查有沒有訂閱電子報
	sqlepaper = "select * from Epaper where email ='"& RSreg("email") &"'"
	set RSepaper = conn.execute(sqlepaper)
	if not RSepaper.eof then
		checkmail = 1
	else
		checkmail = 0
	end if 
	'檢查有沒有訂閱電子報end
	
	
	
	'檢查有沒有開啟動態游標
	if RSreg("ShowCursorIcon") = 1 then
		checkcursor = 1
	else
		checkcursor = 0
	end if 
	'檢查有沒有開啟動態游標end
	
	
	'備註欄
	if len(RSreg("remark")) > 0 then
		Remark = RSreg("remark")
	else
		Remark = ""
	end if
	
	
	
  Set xup = Server.CreateObject("TABS.Upload")
   
  dim test : test = "" 
   
  if RSreg("id_type2")="1" and RSreg("id_type1")="1" then
		test = "2"
  else
    test = "1"
  end if
   
	if  RSreg.eof then
      response.write "<script>alert('找不到資料');history.back();</script>"
      response.end
	else
	
	
	
	
%>
<script language="JavaScript" type="text/javascript">
<!--
function Send()
{
var rege = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9])+$/;

    if (document.Form1.accounttext.value =="" ) {
		alert("您忘了填寫帳號了！"); 
		document.Form1.accounttext.focus(); 
	return false; 
	} 
	if(document.Form1.accounttext.value.length > 30  ) {
		alert("您所填寫的帳號超過30碼！"); 
		document.Form1.accounttext.focus(); 
	     return false; 
	 }
	if (document.Form1.accounttext.value.length < 6){
        alert("您所填寫的帳號少於6碼！"); 
	    document.Form1.accounttext.focus();
		return false; 
       }	
	if (document.Form1.accounttext.value!=""){
       
	      var i;
		  var ch;
		  for (i=0;i< document.Form1.accounttext.value.length;i++){
			ch = document.Form1.accounttext.value.charAt(i);
			if(ch >= 'a' && ch <= 'z'){
				//return true;
			}
			else if(ch >= 'A' && ch <= 'Z'){
				//return true;
			}
			else if(ch >= '0' && ch <= '9'){
				//return true;
			}
			else if(ch == '-' || ch == '_'){
				//return true;
			}
			else{
				alert("帳號限用英文與數字，可用『-』或『_』！"); 
				return false;
			}
	      }
	}	  
	   
	   
	   
	if (document.Form1.realnametext.value == ""  ) {
		alert("請輸入姓名！"); 
		document.Form1.realnametext.focus(); 
	return false; 
	}  
	if (document.Form1.passwd.value =="") {
	        alert("您忘了填寫密碼了！"); 
			document.Form1.passwd.focus(); 
			return false; 
		}
	if(document.Form1.passwd.value!=""){
            		
		            if (document.Form1.passwd.value.length > 16){
			            alert("您所填寫的密碼超過16碼！"); 
			            document.Form1.passwd.focus(); 
			             return false; 
		            }
		            else if(document.Form1.passwd.value.length < 6){
			            alert("您所填寫的密碼少於6碼！"); 
			            document.Form1.passwd.focus(); 
			            return false; 
		            }
		            else if(document.Form1.password2.value==""){
			            alert("您忘了填寫密碼確認了！"); 
			            document.Form1.password2.focus(); 
			            return false; 
		            }
		            else if(document.Form1.passwd.value != document.Form1.password2.value){
			            alert("密碼與密碼確認不符！");
			            document.Form1.password2.focus(); 
			            return false; 
		            }
		            else{
			            var i;
			            var ch;
			            var digits = "0123456789";
			            var digitflag = false;
			            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
			            var charflag = false;
			            for (i=0;i< document.Form1.passwd.value.length;i++){		
				            ch = document.Form1.passwd.value.charAt(i);		
				            if(ch == '\"' || ch == ' ' || ch == '\'' || ch== '\&'){
					            alert('密碼請勿包含『"』、『單引號』、『&』或空白'); 
					            document.Form1.passwd.focus();
					            return false; 
				            }
				            if( digits.indexOf(ch) >= 0 ) {
					            digitflag = true;
				            }					
				            if( chars.indexOf(ch) >= 0 ) {
					            charflag = true;
				            }
			            }
			            if( !digitflag ) {
					            alert("密碼請至少包含一數字"); 
					            document.Form1.passwd.focus();
					            return false; 
			            }
			            if( !charflag ) {
					            alert("密碼請至少包含一英文字"); 
					           document.Form1.passwd.focus();
					           return false; 
			            }
		            }	
	
	
	}
	 if (document.Form1.emailtext.value == ""  ) {
		alert("請輸入E-mail！"); 
		document.Form1.emailtext.focus(); 
	return false; 
	} 
	
     if (rege.exec(document.Form1.emailtext.value) == null) {
		alert("eMail 格式錯誤！"); 
		document.Form1.emailtext.focus(); 
		return false; 
	}
	if(document.Form1.identity.value == 1 )
	{
	//if (document.Form1.idntext.value == "" ) {
	//	alert("請輸入身分證字號！"); 
	//	document.Form1.idntext.focus(); 
	//return false; 
	//}
	}	
	if(document.Form1.identity.value == 2 )
	{
	if (document.Form1.member_orgtext.value == ""  ) {
		alert("請輸入所屬機關名稱！"); 
		document.Form1.member_orgtext.focus(); 
	return false; 
	}
	 if (document.Form1.com_ext.value.length > 4  ) {
		alert("您所填寫所屬機關電話分機超過4碼！"); 
		document.Form1.com_ext.focus(); 
	return false; 
	} 
	 if (document.Form1.home_exttext.value.length > 4  ) {
		alert("您所填寫電話分機超過4碼！"); 
		document.Form1.home_exttext.focus(); 
	return false; 
	} 
     //if (document.Form1.com_teltext.value == ""  ) {
		//alert("請輸入所屬機關電話！"); 
		//document.Form1.com_teltext.focus(); 
	//return false; 
	//} 
	//else if (document.Form1.com_exttext.value == ""  ) {
		//alert("請輸入所屬機關分機！"); 
		//document.Form1.com_exttext.focus(); 
	//return false; 
	//} 
    if (document.Form1.ptitle.value == ""  ) {
		alert("請輸入職稱！"); 
		document.Form1.ptitle.focus(); 
	return false; 
	} 
	}
	
	
	return true;
	
}


function Check() {
    if ( document.getElementById("gardenPro").checked == true ) {
	document.getElementById("garden_intro").style.display = "inline" ;
	document.getElementById("garden_order").style.display = "inline";
	document.getElementById("keyWord").style.display = "inline";
	document.getElementById("expertImg").style.display = "inline";
	document.getElementById("image").style.display = "inline";
	}
	else {
	document.getElementById("garden_intro").style.display = "none" ;
	document.getElementById("garden_order").style.display = "none";
	document.getElementById("keyWord").style.display = "none";
	document.getElementById("expertImg").style.display = "none";
	document.getElementById("image").style.display = "none";
	}
}
		
-->
</script>
<title>
	管理會員資料
</title>
<link href="../css/layout.css" rel="stylesheet" type="text/css" />
<link href="../css/form.css" rel="stylesheet" type="text/css" /><meta http-equiv="Content-Type" content="text/html; charset=utf-8" /></head>
<% if test = "2" then %>
<body onload="showbutton(2)">
<% else test="1" %>
<body onload="showbutton(1)">
<% end if %>
    <div id="FuncName">
	    <h1>會員資料管理</h1>
		<div id="Nav">
	    <a href="newMemberList.asp">回條列</a>
			<a href="/KpiSetting/KpiCalculateTotal.asp?memberid=<%=id%>">KPI記錄</a>
		</div>
		<div id="ClearFloat"></div>
    </div>
    <div id="FormName">&nbsp;<font size="2">【管理會員資料】</font>    </div>
    <form name="Form1" method="post"  id="Form1" action="newMemberEdit_Act.asp?account=<%=trim(RSreg("account"))%>&nowPage=<%=request.QueryString("nowPage")%>&pagesize=<%=request.QueryString("pagesize")%>&from=<%=request.QueryString("from")%>" onsubmit="return Send()" ENCTYPE="MULTIPART/FORM-DATA" >
		<input type="hidden" name="status" value="<%=RSreg("status")%>" />
		<INPUT TYPE="hidden" name="submitTask" value="" />
		<INPUT TYPE="hidden" name="identity" id="identity" value="" />
		<div>
		</div>

        <div>    
          <table width="90%" cellspacing="0">
          <tr>
            <th colspan="2">會員資料管理</th>
            </tr>
          <tr>
            <td class="Label" align="right"><span class="Must">*</span>會員帳號：</td>
            <td class="eTableContent"><input name="accounttext" readonly type="text" id="accounttext" value="<%=trim(RSreg("account"))%>"  />
			限用英文與數字，可用『-』或『_』，30碼以下</td>
          </tr>
          <tr>
            <td class="Label" align="right"><span class="Must">*</span>真實姓名：</td>
            <td class="eTableContent"><input name="realnametext" type="text" id="realnametext" value="<%=trim(RSreg("realname"))%>" /></td>
          </tr>
          <tr>
            <td class="Label" align="right">暱　　稱：</td>
            <td class="eTableContent"><input name="nicknametext" type="text" value="<%=RSreg("nickname")%>" id="nicknametext" /></td>
          </tr>
          <tr>
            <td class="Label" align="right"><span class="Must">*</span>設定密碼：</td>
            <td class="eTableContent"><input name="passwd" type="password" id="passwd" value="<%=Trim(RSreg("passwd"))%>" />
            
            自訂英文（區分大小寫）、數字，不含空白及@，16碼以下</td>
          </tr>
          <tr>
            <td class="Label" align="right"><span class="Must">*</span>密碼確認：</td>
            <td class="eTableContent"><input name="password2" type="password" id="password2" value="<%=Trim(RSreg("passwd"))%>" /></td>
          </tr>                                    
          <tr>
            <td class="Label" align="right">身分證字號：</td>
            <td class="eTableContent"><input name="idntext" type="password" id="idntext" value="<%=RSreg("id")%>" /><font color="white"><%=RSreg("id")%></font></td>
          </tr>
          <tr>
            <td class="Label" align="right"><div id="orgnamediv" style="display:none"><span class="Must">*</span>所屬機關名稱：</div></td>
            <td class="eTableContent"><div id="orgtextdiv" style="display:none"><input name="member_orgtext" type="text" value="<%=RSreg("member_org")%>" id="member_orgtext" /></div></td>
          </tr>
		  <tr>
            <td class="Label" align="right"><div id="com_telnamediv" style="display:none"><span class="Must">*</span>所屬機關電話：</div></td> 
			<td class="eTableContent">
                <div id="com_teltextdiv" style="display:none"><input name="com_tel" type="text" value="<%=RSreg("com_tel")%>" id="com_tel"  /></div></tr>
          <tr>
            <td class="Label" align="right"><div id="com_extnamediv" style="display:none"><label for="com_ext">分機：</label></div>
            <td><div id="com_exttextdiv" style="display:none"><input name="com_ext" type="text" id="com_ext" value="<%=RSreg("com_ext")%>" /> </div></td>
				
          </tr>
          <tr>
            <td class="Label" align="right"> <div id="ptitlenamediv" style="display:none"><span class="Must">*</span>職稱：</div></td>
            <td class="eTableContent"><div id="ptitletextdiv" style="display:none"><input name="ptitle" type="text" value="<%=RSreg("ptitle")%>" id="ptitle" /></div></td>
          </tr>
          <tr>
            <td class="Label" align="right">出生日期：</td>      
            <td class="eTableContent">
<% 
dim birthyear : birthyear = ""
if not IsNULL( RSreg("birthday") ) then 
	birthyear = Mid(RSreg("birthday"),1,4)
end if
dim birthmonth : birthmonth =""
if not IsNULL( RSreg("birthday") ) then 
	birthmonth = Mid(RSreg("birthday"),5,2)
end if
dim birthdaytext : birthdaytext =""
if not IsNULL( RSreg("birthday") ) then 
	birthdaytext = Mid(RSreg("birthday"),7,2)
end if

%>
        西元<input name="birthYeartext" type="text" value="<%=birthyear%>" id="birthYeartext" />年
				
                <select name="birthMonthtext" id="birthMonthtext">
    <option value=""  <%if birthmonth=""   then %> selected<%end if%>>請選擇</option>
	<option value="01"<%if birthmonth="01" then %> selected<%end if%>>1</option>
	<option value="02"<%if birthmonth="02" then %> selected<%end if%>>2</option>
	<option value="03"<%if birthmonth="03" then %> selected<%end if%>>3</option>
	<option value="04"<%if birthmonth="04" then %> selected<%end if%>>4</option>
	<option value="05"<%if birthmonth="05" then %> selected<%end if%>>5</option>
	<option value="06"<%if birthmonth="06" then %> selected<%end if%>>6</option>
	<option value="07"<%if birthmonth="07" then %>selected<%end if%>>7</option>
	<option value="08"<%if birthmonth="08" then %> selected<%end if%>>8</option>
	<option value="09"<%if birthmonth="09" then %>selected<%end if%>>9</option>
	<option value="10"<%if birthmonth="10" then %> selected<%end if%>>10</option>
	<option value="11"<%if birthmonth="11" then %>selected<%end if%>>11</option>
	<option value="12"<%if birthmonth="12" then %>  selected<%end if%>>12</option>
</select>
                月
                <select name="birthdaytext" id="birthdaytext">
	<option value=""   <%if birthmonth=""     then%>selected<%end if%>>請選擇</option>	
	<option value="01" <%if birthdaytext="01" then%>selected<%END IF%>>1</option>
	<option value="02" <%if birthdaytext="02" then%>selected<%END IF%>>2</option>
	<option value="03" <%if birthdaytext="03" then%>selected<%END IF%>>3</option>
	<option value="04" <%if birthdaytext="04" then%>selected<%END IF%>>4</option>
	<option value="05" <%if birthdaytext="05" then%>selected<%END IF%>>5</option>
	<option value="06" <%if birthdaytext="06" then%>selected<%END IF%>>6</option>
	<option value="07" <%if birthdaytext="07" then%>selected<%END IF%>>7</option>
	<option value="08" <%if birthdaytext="08" then%>selected<%END IF%>>8</option>
	<option value="09" <%if birthdaytext="09" then%>selected<%END IF%>>9</option>
	<option value="10" <%if birthdaytext="10" then%>selected<%END IF%>>10</option>
	<option value="11" <%if birthdaytext="11" then%>selected<%END IF%>>11</option>
	<option value="12" <%if birthdaytext="12" then%>selected<%END IF%>>12</option>
	<option value="13" <%if birthdaytext="13" then%>selected<%END IF%>>13</option>
	<option value="14" <%if birthdaytext="14" then%>selected<%END IF%>>14</option>
	<option value="15" <%if birthdaytext="15" then%>selected<%END IF%>>15</option>
	<option value="16" <%if birthdaytext="16" then%>selected<%END IF%>>16</option>
	<option value="17" <%if birthdaytext="17" then%>selected<%END IF%>>17</option>
	<option value="18" <%if birthdaytext="18" then%>selected<%END IF%>>18</option>
	<option value="19" <%if birthdaytext="19" then%>selected<%END IF%>>19</option>
	<option value="20" <%if birthdaytext="20" then%>selected<%END IF%>>20</option>
	<option value="21" <%if birthdaytext="21" then%>selected<%END IF%>>21</option>
	<option value="22" <%if birthdaytext="22" then%>selected<%END IF%>>22</option>
	<option value="23" <%if birthdaytext="23" then%>selected<%END IF%>>23</option>
	<option value="24" <%if birthdaytext="24" then%>selected<%END IF%>>24</option>
	<option value="25" <%if birthdaytext="25" then%>selected<%END IF%>>25</option>
	<option value="26" <%if birthdaytext="26" then%>selected<%END IF%>>26</option>
	<option value="27" <%if birthdaytext="27" then%>selected<%END IF%>>27</option>
	<option value="28" <%if birthdaytext="28" then%>selected<%END IF%>>28</option>
	<option value="29" <%if birthdaytext="29" then%>selected<%END IF%>>29</option>
	<option value="30" <%if birthdaytext="30" then%>selected<%END IF%>>30</option>
	<option value="31" <%if birthdaytext="31" then%>selected<%END IF%>>31</option>
</select>       
                日			</td>
          </tr>
		
          <tr>
            <td class="Label" align="right">性別：</td>                  
            <td class="eTableContent" align="left">
             <input id="maletext" type="radio" name="sex"value="1" <%if trim(RSreg("sex"))="1" then%> checked <%END IF%> /><label for="maletext">男</label>
				
	<input id="femaletext" type="radio" name="sex"  value="0"<%if trim(RSreg("sex"))="0" then%> checked <%END IF%>  />
				<label for="femaletext">女</label>	</td>
          </tr>
		  
		
          <tr>
            <td class="Label" align="right">地址：</td>               
            <td class="eTableContent">
			    <input name="homeaddrtext" type="text" value="<%=Trim(RSreg("homeaddr"))%>"  id="homeaddrtext" />
                <label for="zip"> 郵遞區號：</label>
                <input name="ziptext" type="text" id="ziptext" value="<%=Trim(RSreg("zip"))%>" />    </td>
          </tr>
          <tr>
            <td class="Label" align="right">電話：</td>              
            <td class="eTableContent">
              <input name="phonetext" type="text" value="<%=Trim(RSreg("phone"))%>"                                                                                                     " id="phonetext" />
              <label for="home_ext">分機：</label>
              <input name="home_exttext" type="text" id="home_exttext" value="<%=Trim(RSreg("home_ext"))%>"  />            </td>
          </tr>
          <tr>
            <td class="Label" align="right">手機號碼：</td>    
            <td class="eTableContent"><input name="mobiletext" type="text" value="<%=Trim(RSreg("mobile"))%>"                                                                                                  " id="mobiletext" /></td>
          </tr>
          <tr>            
            <td class="Label" align="right">傳真：</td>    
            <td class="eTableContent"><input name="faxtext" type="text" id="faxtext" value="<%=Trim(RSreg("fax"))%>" /></td>
          </tr>
          <tr>            
            <td class="Label" align="right"><span class="Must">*</span>E-mail：</td>    
            <td class="eTableContent"><input name="emailtext" type="text" value="<%=Trim(RSreg("email"))%>" id="emailtext" />
			供系統認證之用，請務必填寫正確</td>
          </tr>
          <tr>            
            <td class="Label" align="right">Email 認證：</td>    
            <td class="eTableContent">
			<% If RSreg("mcode") = "Y" Then %>
			<INPUT TYPE="radio" NAME="emailverified" VALUE="Y" CHECKED />已通過&nbsp;<INPUT TYPE="radio" NAME="emailverified" VALUE="" />未通過&nbsp;
			<% Else %>
			<INPUT TYPE="radio" NAME="emailverified" VALUE="Y"/>已通過&nbsp;<INPUT TYPE="radio" NAME="emailverified" VALUE="" CHECKED />未通過&nbsp;
			<% End If %>&nbsp;
			(申請驗證次數：<%=RSreg("ValidCount") %>)
			</td>
          </tr>
		  <tr>
            <td class="Label" align="right">電子報訂閱：</td>                  
            <td class="eTableContent" align="left">
             <input id="epapercheck" type="checkbox" name="epapercheck" value="1" <%if checkmail = 1 then%> checked <%END IF%> />訂閱電子報
			 &nbsp;&nbsp;&nbsp;&nbsp;<font  style="color:#005555;">是否開啟動態游標：</font>
			 <input id="cursorcheck" type="checkbox" name="cursorcheck" value="1" <%if checkcursor = 1 then%> checked <%END IF%> />開啟動態游標 
			</td>
		                
      
		  <tr>
			   <td class="Label" align="right">備註：</td>	
			   <td><textarea name="remark" cols="50" rows="4" id="remark" class="InputText"><%=Remark%></textarea></td>
          </tr>
          <tr>
            <td class="Label" align="right"><div id="KMcatnamediv" style="display:none">研究領域及專長：</div></td>
            <td> <div id="KMcattextdiv" style="display:none"><input class="Button" type="text" size="30" name="KMcat"  id="KMcat" value="<%=trim(RSreg("KMcat"))%>" ></div>
						<input class="Text" type="hidden" size="60" name="htx_KMcatID"  id="htx_KMcatID" >
						<input class="Text" type="hidden" size="60" name="htx_KMautoID"  id="htx_KMautoID" >
						 <!--<BUTTON  id="btn_KMcat" class="Text" >外部分類</BUTTON>/-->
			
                            </td>
          </tr>
          </table>
          <hr />
          <table width="90%" cellspacing="0">
          <tr>
            <th colspan="2">會員權限管理</th>
          </tr>
          <tr>
            <td class="Label" align="right">上傳圖片權限：</td>
            <td class="eTableContent"> 目前狀態為          				
              <select name="uploadRight" size="1">
                <option value=""  <% if trim(RSreg("uploadRight"))="" THEN %>selected <%END IF%>>請選擇</option>
                <option value="Y" <% if trim(RSreg("uploadRight"))="Y" THEN %>selected <%END IF%>>正常</option>
                <option value="N" <% if trim(RSreg("uploadRight"))="N" THEN %>selected <%END IF%>>停權</option>
              </select>
							可上傳圖片數量/ 每次<input name="uploadPicCount" type="text" id="uploadPicCount" value="<%=RSreg("uploadPicCount")%>" size="10" />張
						</td>
          </tr>
          <tr>
            <td align="right" class="Label">會員身分管理：</td>
            <td class="eTableContent">一般會員 / 學者會員
							<font color="red">
								<% 
									if RSreg("scholarValidate")="Z" then 
										response.write "(無)"
									elseif RSreg("scholarValidate") = "W" then 
										response.write "(待審核)"
									elseif RSreg("scholarValidate") = "Y" then 
										response.write "(通過)"
									elseif RSreg("scholarValidate") = "N" then 
										response.write "(不通過)"
									end if
								%>
							</font>
							<% if RSreg("scholarValidate") = "W" Or RSreg("scholarValidate") = "Y" Or RSreg("scholarValidate") = "N" then %>
							<input type="button" name="ScholarPassBtn" value="通過" onclick="ScholarPass()" id="ScholarPassBtn" class="cbutton" />  
							<input type="button" name="ScholarNotPassBtn" value="不通過" onclick="ScholarNotPass()" id="ScholarNotPassBtn" class="cbutton" />  
							<% end if %>
						</td>
          </tr>
		  <!--新增功能:可直接改為學者會員身分-->
		  <%if RSreg("scholarValidate") = "Z" then%>
		    <tr>
		      <td align="right" class="Label">學者會員：</td>
			  <td>
			  <input id="scholarcheck" type="checkbox" name="scholarcheck" value="1" />改為學者會員身份
              </td>		
		    </tr>
		  <%end if%>
          </table>
          <hr />
          <table width="90%" cellspacing="0">
            <tr>
              <th colspan="2">專家身分管理</th>
            </tr>
            <tr>
              <td class="Label" align="right">專家：</td>			  
              <td class="eTableContent">
				<input name="knowledgePro" type="checkbox" value="1"  <% if trim(RSreg("id_type3"))="1" THEN %>  checked   <%END IF%> />
                <span class="Label">設為知識家專家</span>			  
				<input name="gardenPro" id="gardenPro" type="checkbox" value="1"  <% if not gardenExpert_RSreg.eof THEN %>  checked  <%END IF%>  onclick='Check()'/>
                 <span class="Label">設為園藝專家</span></td>
            </tr>
						
			<tr id='keyWord'>
				<td class="Label" align="right">專長關鍵字：</td>
				<td class="eTableContent">
					<input name="keyword" class="rdonly" title="請以,分隔" value="<%=RSreg("keyword")%>" size="75" readonly="true" /><br/>
					<input name="button2" type="button" class="cbutton" onclick="setKeywords()" value ="設定" />              
				</td>
			</tr>
			
            <tr id='expertImg'>
              <td class="Label" align="right">專家圖檔：</td>			
              <td class="eTableContent"><input type="file" name="imgfile" id="imgfile" /></td>
		    </tr>
			<tr id='image'>
			  <td></td>
			  <%if RSreg("photo")<>"" then%>
			  <td class="eTableContent"><img src="<%=RSreg("photo")%>" align="center" width="150" height="100" border="0"></td>
			  <%end if %>
            </tr>
			<tr id ='garden_intro' >
			   <td class="Label" align="right">自我介紹：</td>	
			   <td><textarea name="textfield" cols="50" rows="4" id="textfield" class="InputText"><%=gardenExpert_intro%></textarea></td>
			</tr>
			<tr  id ='garden_order' >
				<td class="Label" align="right">園藝專家排序：</td>	
				<td><input name="order" type="text" size="2" id="order" value="<%=gardenExpert_order%>"/></td>
			</tr>
			<% if gardenExpert_RSreg.eof THEN %>
			<script language="javascript">
			document.getElementById("garden_intro").style.display = "none" ;
	        document.getElementById("garden_order").style.display = "none";
           	document.getElementById("keyWord").style.display = "none";
	        document.getElementById("expertImg").style.display = "none";
	        document.getElementById("image").style.display = "none";
			</script>
			<%END IF%>
          </table>
            <hr />
            <table width="90%" cellspacing="0">
                <tr>
                    <th colspan="2">已邀請成功的朋友清單</th>
                </tr>
                <tr>
                    <td style="width:200px"></td>
                    <td align="center">
                        <table style="margin:5px 0px 0px 20px ; width:450px; border-color:#c6c6c6" id="InviteeTable" border="1" cellpadding="1" cellspacing="0">
                            <tr id="header">
                                <td>&nbsp;&nbsp;</td>
                                <td>朋友帳號</td>
                                <td>朋友暱稱</td>
                                <td>註冊時間</td>
                                <td>通過認證</td>
                            </tr>
                            <%
                             idx=0                             
                             invitationStr = ""
                             invitationStr = invitationStr & vbcrlf & "SELECT "
                             invitationStr = invitationStr & vbcrlf & "	 M.nickname"
                             invitationStr = invitationStr & vbcrlf & "	,M.account "
                             invitationStr = invitationStr & vbcrlf & "	,M.createtime "
                             invitationStr = invitationStr & vbcrlf & "	,I.IsActive "
                             invitationStr = invitationStr & vbcrlf & " FROM    member AS M "
                             invitationStr = invitationStr & vbcrlf & " INNER JOIN InviteFriends_Detail AS I ON M.account = I.inviteAccount "
                             invitationStr = invitationStr & vbcrlf & " WHERE  I.InvitationCode = (select InvitationCode from InviteFriends_Head where account='" & id & "')"
                             
                             set rs = conn.execute(invitationStr)
                             if rs.eof then
                                response.Write "<tr>"
                                response.Write "<td colspan='10'>尚無邀請好友</td>"
                                response.Write "</tr>"                             
                             end if
                             do while not rs.eof
                                idx = idx+1
                                response.Write "<tr>"
                                response.Write "<td>" & idx & "</td>"
                                response.Write "<td>" & rs("account") & "</td>"
                                response.Write "<td>" & rs("nickname") & "</td>"
                                response.Write "<td>" & rs("createtime") & "</td>"
                                
                                select case cstr(rs("IsActive"))
                                    case "0"
                                        response.Write "<td>否</td>"
                                    case "1"
                                        response.Write "<td>是</td>"
                                    case "9"
                                        response.Write "<td>停權</td>"
                                end select
                                
                                response.Write "</tr>"
                                rs.movenext
                             loop
                            %>                            
                        </table>
                    </td>
                </tr>
            </table>
            <%
%>
		   <input type="submit" name="submit" value ="編修存檔" class="cbutton" >
		   <% IF RSreg("status")="Y"  then%>
           <input type="button" name="ResetBtn" value="停權" onclick="Suspension()" id="ResetBtn" class="cbutton" />   
		   <%else %>
		   <input type="button" name="ResetBtn" value="復權" onclick="RecoverRight()" id="ResetBtn" class="cbutton" /> 	
           <%end if %>		   
           <input type="button" name="ResetBtn" class="cbutton" value="取消" onclick="returnToEdit()" id="ResetBtn" class="cbutton" />   
           
           

        </div>
    
<div>

</div>
</form>
</body>
</html>
<%  end if%>
<script language="javascript">
	
	function setKeywords() 
	{
		window.location.href = "keywords_set.asp?memberId=<%=id%>&nowPage=<%=request("nowPage")%>&pagesize=<%=request("pagesize")%>";
  }
	function returnToEdit() 
	{
	    if (getParameter("from") == "query") {
			window.location.href = "newMemberQuery_Act.asp?nowPage=<%=request("nowPage")%>&pagesize=<%=request("pagesize")%>";
		}
		else {
		window.location.href = "newMemberList.asp?nowPage=<%=request("nowPage")%>&pagesize=<%=request("pagesize")%>";
  }
    }
	function getParameter(parameterName) {  
        var url = location.search.substring(1);  
        parameterName = parameterName + "=";  
        if (url.length > 0) {
            begin = url.indexOf(parameterName);
            if (begin != -1) {
                begin += parameterName.length;
                end = url.indexOf("&" , begin);  
                if ( end == -1 ) end = url.length  
                return unescape(url.substring(begin, end));  
            }  
            return "null";  
       }  
   }  
	function  Suspension(){	 
		window.location.href = "newMemberEdit.asp?account=<%=request("account")%>&Status=Suspension&nowPage=<%=request("nowPage")%>&pagesize=<%=request("pagesize")%>";	 
	}
	function ScholarPass() 
	{
		var rege = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9])+$/;
		if( document.getElementById("emailtext").value == "") {
			alert("請輸入email");
		}
		else if( rege.exec(document.Form1.emailtext.value) == null ) {
			alert("email 格式錯誤！"); 
		}
		else {
			window.location.href = "newMemberEdit.asp?account=<%=request("account")%>&scholar=pass&nowPage=<%=request("nowPage")%>&pagesize=<%=request("pagesize")%>&email=" + document.getElementById("emailtext").value;	 
		}		
	}
	function ScholarNotPass() 
	{
		var rege = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9])+$/;
		if( document.getElementById("emailtext").value == "") {
			alert("請輸入email");
		}
		else if( rege.exec(document.Form1.emailtext.value) == null ) {
			alert("email 格式錯誤！"); 
		}
		else {
			window.location.href = "newMemberEdit.asp?account=<%=request("account")%>&scholar=notpass&nowPage=<%=request("nowPage")%>&pagesize=<%=request("pagesize")%>&email=" + document.getElementById("emailtext").value;	 
		}	
	}
	function  RecoverRight()
	{
		window.location.href = "newMemberEdit.asp?account=<%=request("account")%>&Status=RecoverRight&nowPage=<%=request("nowPage")%>&pagesize=<%=request("pagesize")%>";	 
	}
	
	function showbutton(id) {
	   if (id == 1) {
			document.getElementById("orgnamediv").style.display = "none";	           	
			document.getElementById("orgtextdiv").style.display = "none";
            document.getElementById("com_telnamediv").style.display = "none";	           	
			document.getElementById("com_teltextdiv").style.display = "none";
			document.getElementById("com_extnamediv").style.display = "none";	           	
			document.getElementById("com_exttextdiv").style.display = "none";
			document.getElementById("ptitlenamediv").style.display = "none";	           	
			document.getElementById("ptitletextdiv").style.display = "none";
			document.getElementById("KMcatnamediv").style.display = "none";	           	
			document.getElementById("KMcattextdiv").style.display = "none";
			document.getElementById("identity").value =1;
			
		}
		if (id == 2) {
			document.getElementById("orgnamediv").style.display = "block";	           	
			document.getElementById("orgtextdiv").style.display = "block";	
            document.getElementById("com_telnamediv").style.display = "block";	           	
			document.getElementById("com_teltextdiv").style.display = "block";
	        document.getElementById("com_extnamediv").style.display = "block";	           	
			document.getElementById("com_exttextdiv").style.display = "block"; 
            document.getElementById("ptitlenamediv").style.display = "block";	           	
			document.getElementById("ptitletextdiv").style.display = "block"; 
            document.getElementById("KMcatnamediv").style.display = "block";	           	
			document.getElementById("KMcattextdiv").style.display = "block"; 
            document.getElementById("identity").value =2;        	
		}
	 }
	                                	                  
	
</script>
<script language="vbs">                     
	'document.domain = "coa.gov.tw"
	'sub btn_KMcat_onClick()
	'	window.open "http://kmintra.coa.gov.tw/coa/ekm/manage_doc/report2cat22.jsp?data_base_id=DB020&id_name=htx_KMcatID&autoid_name=htx_KMautoID&nm_name=htx_KMcat&&subNode=1*RB1&display=1*none&form_name=form1&focus=&catidsInput=&anchor=",null,"height=400,width=500,status=no,scrollbars=yes,toolbar=no,menubar=no,location=no"
	'end sub
</script>

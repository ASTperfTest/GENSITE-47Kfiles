<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<% 
on error resume next
response.expires = 0 
%>
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<%
	'改回舊的連線
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open session("ODBCDSN")
	'改回舊的連線end
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
	Response.Buffer = true
	Dim Message

	'身份證字號不用填

	if Request("account") <> "" and Request("passwd") <> "" and Request("realname") <> ""  and Request("email") <> ""  Then 

		'判斷帳號是否已使用
		set rs = conn.Execute("select * from Member Where account ='" & Request("account") & "'")
		if rs.eof then
			'判斷是否已是會員
			Set rsid = conn.Execute("select * from Member Where id ='" & Request("idn") & "'")
			
			'判斷是否有填寫出生年
			if Request("birthYear") <>"" then
				Dim xbirtmm
				Dim xbirtdd
				Dim xbirthday
				if Request("birthMonth") <=9 then 
					xbirtmm ="0" & Request("birthMonth")
				else
					xbirtmm = Request("birthMonth")
				end if
				
				if Request("birthday") <=9 then 
					xbirtdd ="0" & Request("birthday")
				else
					xbirtdd = Request("birthday")
				end if
			
				xbirthday = Request("birthYear") & xbirtmm & xbirtdd
			else
				xbirthday=""
			end if
			
			if rsid.eof then
      	'新增帳號
				Dim orderepaper : orderepaper = "Y"
				if request("orderepaper") <> "Y" then orderepaper = "N"
					
        sql = "INSERT INTO Member( account, passwd, realname, homeaddr, phone, mobile, email, " & _ 
							"createtime, modifytime, zip, home_ext, birthday, sex, id_type1, fax, create_user, id, " & _ 
							"nickname, actor, status, orderepaper, scholarValidate, uploadRight, uploadPicCount ) VALUES( " & pkStr(Request("account"),"") & ", " & pkStr(Request("passwd"),"") & ", " & _
							"" & Chg_UNI(pkStr(Request("realname"),"")) & ", " & pkStr(Request("homeaddr"),"") & ", " & pkStr(Request("phone"),"") & ", " & _
							"" & pkStr(Request("mobile"),"") & ", " & pkStr(Request("email"),"") & ", getdate(), getdate(), " & pkStr(Request("zip"),"") & ", " & _
							"" & pkStr(Request("home_ext"),"") & ",'" & xbirthday & "', " & pkStr(Request("sex"),"") & ", '1', "	& pkStr(Request("fax"),"") & ", " & _
							"'hyweb', " & pkStr(Request("idn"),"") & ", " & Chg_UNI(pkStr(Request("nickname"),"")) & ", '0', 'Y', '" & orderepaper & "', 'Z', 'Y', 1)"
				'response.write(sql) & Request("account")
				conn.Execute(sql)

				'訂閱電子報
				if request("orderepaper") = "Y" then
					Dim emailFlag
					emailFlag = true
					'檢查是否已經訂閱
					set ts = conn.Execute("select count(*) from Epaper where email = '" & Request("email") & "' and CtRootID = 21")
					If Err.Number <> 0 or conn.Errors.Count <> 0 Then
						emailFlag = false
					end if		  
					if ts(0) > 0 then
						emailFlag = false
					end if
	  
					'檢查email格式是否正確
					if len(Request("email")) > 3 and InStr(Request("email"), "@") > 0 and InStr(Request("email"), ".") > 0 and emailFlag = true then
						'正確...存入資料庫
						sql = "insert into Epaper ( email, createtime, CtRootID) values ('"& Request("email") & "', getdate(), 21)"
						conn.execute(sql)	    
					end if
				end if
				if err.number <> 0 then
					Response.Write "<html><body bgcolor='#ffffff'>"
					Response.Write "<script language='javascript'>"
					Response.Write "alert('申請會員發生錯誤,請洽系統管理員');"
					Response.Write "location.replace('sp.asp?xdURL=coamember/member_CompleteC.asp');"
					'Response.Write "location.replace('member_CompleteC.asp')"
					Response.Write "</script>"
					Response.Write "</body></html>" 
				else
					'寄出通知信  
					response.write "<HR> 寄出通知信"
					'response.end         
					%>
					<!--#Include file = "mailbody_C.inc" -->
					<%
					Response.Write "<html><body bgcolor='#ffffff'>"
					Response.Write "<script language='javascript'>"
					Response.Write "alert('我們將以E-mail通知您的申請狀況!');"
					Response.Write "location.replace('sp.asp?xdURL=coamember/member_CompleteC.asp');"
					'Response.Write "location.replace('member_CompleteC.asp')"
					Response.Write "</script>"
					Response.Write "</body></html>" 
				end if  				
			else 
				'如果id已存在提示已具有會員身份
				Response.Write "<html><body bgcolor='#ffffff'>"
				Response.Write "<script language='javascript'>"
				Response.Write "alert('您已具有會員身份!!');"
				Response.Write "history.back();"
				'Response.Write "location.replace('/index.jsp');"
				'Response.Write "location.replace('/');"
				Response.Write "</script>"
				Response.Write "</body></html>"  
			end if
			'response.end  			
    else
			'如果account已存在提示改用其他帳號
			Response.Write "<html><body bgcolor='#ffffff'>"
			Response.Write "<script language='javascript'>"
			Response.Write "alert('此帳號已經存在, 請改用其他帳號!');"
			Response.Write "history.back();"
			Response.Write "</script>"
			Response.Write "</body></html>"   
   	end if
	else
		Response.Write "<html><body bgcolor='#ffffff'>"
		Response.Write "<script language='javascript'>"
		Response.Write "alert('請填入必要資料以供審核, 謝謝!');"
		Response.Write "history.back();"
		Response.Write "</script>"
		Response.Write "</body></html>"  
	end if
	
Function Chg_UNI(str)        'ASCII轉Unicode
	dim old,new_w,iStr
	old = str
	new_w = ""
	for iStr = 1 to len(str)
		if ascw(mid(old,iStr,1)) < 0 then
			new_w = new_w & "&#" & ascw(mid(old,iStr,1))+65536 & ";"
		elseif        ascw(mid(old,iStr,1))>0 and ascw(mid(old,iStr,1))<127 then
			new_w = new_w & mid(old,iStr,1)
		else
			new_w = new_w & "&#" & ascw(mid(old,iStr,1)) & ";"
		end if
	next
	Chg_UNI=new_w
End Function
%>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<% 
on error resume next
response.expires = 0 
%>
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
	Response.Buffer = true
	
	If Session("memID") <> "" Then

		if Request("account") <> "" and Request("passwd1") <> "" and Request("email") <> "" Then 

			'---vincent:判斷傳入的會員帳密是否正確---
			set rs = conn.Execute("select * from Member Where account = " & pkStr( Request("account"), "") & " AND passwd = " & pkStr(request("passwd1"), "") & " ")
		
			if rs.eof then
				'---傳入帳密錯誤,導回首頁---
				response.write "<script>alert('原密碼輸入錯誤!!');history.back();</script>"
			
			else
				'---傳入帳密正確---
			
				'---判斷是否有填寫出生年---
				if Request("birthYear") <> "" then
					Dim xbirtmm
					Dim xbirtdd
					Dim xbirthday
					if Request("birthMonth") <= 9 then 
						xbirtmm = "0" & Request("birthMonth")
					else
						xbirtmm = Request("birthMonth")
					end if
				
					if Request("birthday") <= 9 then 
						xbirtdd = "0" & Request("birthday")
					else
						xbirtdd = Request("birthday")
					end if			
					xbirthday = pkStr(Request("birthYear") & xbirtmm & xbirtdd, "")
				else
					xbirthday = pkStr("", "")
				end if
			
				'---判斷actor---

				Dim account, passwd, nickname, sex, addr, zip, phone, home_ext, mobile, fax, email
				Dim actor, member_org, com_tel, com_ext, ptitle
							
				account = pkStr(Request("account"), "")
				if Request("passwd2") <> "" Then
					passwd = pkStr(Request("passwd2"), "")
				else
					passwd = pkStr(Request("passwd1"), "")
				end if
				nickname = pkStr(Request("nickname"), "")
				sex = pkStr(Request("sex"), "")
				addr = pkStr(Request("addr"), "")
				zip = pkStr(Request("zip"), "")
				phone = pkStr(Request("phone"), "")
				home_ext = pkStr(Request("home_ext"), "")
				mobile = pkStr(Request("mobile"), "")
				fax = pkStr(Request("fax"), "")
				email = pkStr(Request("email"), "")
							
				If Request("actor") = "1" Or Request("actor") = "2" Or Request("actor") = "3" Then
	      	
					actor = pkStr(Request("actor"), "")
					member_org = pkStr(Request("member_org"), "")
					com_tel = pkStr(Request("com_tel"), "")
					com_ext = pkStr(Request("com_ext"), "")
					ptitle = pkStr(Request("ptitle"), "")
					
					sql = "UPDATE Member SET passwd = " & passwd & ", nickname = " & nickname & ", birthday = " & xbirthday & ", " & _
								"sex = " & sex & ", homeaddr = " & addr & ", zip = " & zip & ", phone = " & phone & ", home_ext = " & home_ext & ", " & _
								"mobile = " & mobile & ", fax = " & fax & ", email = " & email & ", actor = " & actor & ", member_org = " & member_org & ", " & _
								"com_tel = " & com_tel & ", com_ext = " & com_ext & ", ptitle = " & ptitle & " WHERE account = " & account	
				Else
				
					sql = "UPDATE Member SET passwd = " & passwd & ", nickname = " & nickname & ", birthday = " & xbirthday & ", " & _
							"sex = " & sex & ", homeaddr = " & addr & ", zip = " & zip & ", phone = " & phone & ", home_ext = " & home_ext & ", " & _
							"mobile = " & mobile & ", fax = " & fax & ", email = " & email & " WHERE account = " & account	
				End If
            
				'response.write(sql)
				conn.Execute(sql)
				
				'訂閱電子報處理
				epapercheck = Request("epapercheck")
				if epapercheck = "1" then
					checksql = "select * from Epaper where email = "& email
					set check_epaper = conn.execute(checksql)
					if check_epaper.eof then
						sql1 = "INSERT INTO Epaper (email,createtime,CtRootID) VALUES ("&email&", getdate(),'21')"
						conn.execute(sql1)
					end if
				else
					sql2 = "DELETE FROM Epaper WHERE email = "& email
					conn.execute(sql2)
				end if
				'訂閱電子報處理end
			
				response.write "<script>alert('修改個人資料成功');window.location.href='" & session("myURL") & "';</script>"			
			end if
		else
			response.write "<script>alert('請輸入正確資料!!');history.back();</script>"
		end if
			
	Else
		response.redirect "../mp.asp?mp=1"
	End if
		
%>

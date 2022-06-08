<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/des.inc" -->
<%
	Dim email,token,returnUrl,updateMemID,key

	email = Request("email")
	token = Request("token")
	returnUrl = Request("returnUrl")
	updateMemID = Request("id")

	key = session("WebDESKey")			'加解密金鑰(Len=14)
	Set DesCrypt = New Cls_DES
	token = DesCrypt.DES(token,key,1)	'解密
	Set DesCrypt = Nothing

	If (token <> email or instr(updateMemID,"'")>0 ) Then	
	    Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head><body bgcolor='#ffffff'>"
	    Response.Write "<script language='javascript'>"
	    Response.Write "alert(' Email 認證失敗！請重新申請或尋求協助！\n" & email & "');"
	    Response.Write "this.close();"
	    Response.Write "</script>"
	    Response.Write "</body></html>"
	    Response.End()		
		
	ElseIf (token = email And returnUrl <> "") Then

		'新會員加入
		session("CheckCodeforMail") = email
		Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head><body bgcolor='#ffffff'>"
		Response.Write "<script language='javascript'>"
		Response.Write "alert(' Email 認證成功！您現在可以繼續流程，開始輸入您的個人資料！');"
		Response.Write "location.replace('" + returnUrl + "');"
		Response.Write "</script>"
  		Response.Write "</body></html>"
		Response.End()

	ElseIf (token = email And updateMemID <> "") Then
		
		'舊會員重新認證
		Dim mcode
		Set rs = conn.Execute("select mcode from Member where email='"& email &"' and (account = '" & updateMemID & "')")	
		if not rs.eof then		
		
			mcode = rs("mcode")
			if IsNull(mcode) then
				conn.Execute("update Member set mcode='Y' where (account = '" & updateMemID & "')") 
				
				
                '增加Kpi 邀請 & 被邀請                
                sqlstr="select * from InviteFriends_head where invitationCode in ("
                sqlstr=sqlstr & vbcrlf &"select invitationCode from InviteFriends_detail where inviteAccount='" & updateMemID & "' and IsActive=0)"
                
                set rsInviter = conn.Execute(sqlstr)
                if not rsInviter.eof then
                    ' 當日日期
                    dateNow = Year(now()) & "/" & Month(now()) & "/" & Day(now())
                    ' 取出邀請者帳號
                    strInviterAccount = trim(rsInviter("account"))
                    
                    invitationCode=rsInviter("invitationCode")
                    
                    ' 取得 邀請 & 被邀請 Kpi贈與分數  ''' Kpi代號
                    set rsInviterScore = conn.Execute("SELECT * FROM kpi_set_score WHERE Rank0_2 = 'st_315'")
                      intInviterScore = trim(rsInviterScore("Rank0_1"))
                    set rsIsInvitedScore = conn.Execute("SELECT * FROM kpi_set_score WHERE Rank0_2 = 'st_316'")
                      intIsInvitedScore = trim(rsIsInvitedScore("Rank0_1"))

                    sno=0
                    ' 檢查邀請者當日是否有登入紀錄
                    set rsInvite = conn.Execute("SELECT * FROM MemberGradeLogin WHERE memberId = '" & strInviterAccount & "' AND convert(varchar, loginDate, 111) = convert(varchar, getdate(), 111) order by sno desc")
                    if not rsInvite.eof then
                        ' 有紀錄，Update當日紀錄
                       sno = trim(rsInvite("sno"))
                       intOriloginInvite = trim(rsInvite("loginInvite"))
                       intNewloginInvite = CInt(intOriloginInvite) + CInt(intInviterScore)
                       strSQLmodiInviterKpi = "UPDATE MemberGradeLogin SET loginInvite = " & _
                                               intNewloginInvite & " WHERE memberId = '" & strInviterAccount & _
                                               "' AND sno = " & sno
                       'response.Write(strSQLmodiInviterKpi & "<br/>")
                       conn.Execute(strSQLmodiInviterKpi)
                    else
                        ' 無紀錄，Insert當日紀錄
                        strSQLaddInviterKpi = "INSERT INTO MemberGradeLogin (memberId, loginInterDate,loginInvite) " & _
                                              "VALUES ('" & strInviterAccount & "', getdate(), " & intInviterScore & ")"
                        'response.Write(strSQLaddInviterKpi & "<br/>")
                        conn.Execute(strSQLaddInviterKpi)
                        
                        set rsInvite = conn.Execute("SELECT sno FROM MemberGradeLogin WHERE memberId = '" & strInviterAccount & "' AND convert(varchar, loginDate, 111) = convert(varchar, getdate(), 111) order by sno desc")
                        sno = trim(rsInvite("sno"))
                    end if

                    ' 增加被邀請者的KPI
                    strSQLaddIsInvited = "INSERT INTO MemberGradeLogin (memberId, loginInterDate,loginIsInvited) " & _
                                              "VALUES ('" & updateMemID & "', getdate(), " & intIsInvitedScore & ")"
                    'response.Write(strSQLaddIsInvited & "<br/>")
                    conn.Execute(strSQLaddIsInvited)
                    
                    
                    strSQLaddIsInvited="update InviteFriends_detail set IsActive=1 , KPIAddedSNO=" & sno & ", Inviter_KPIPointe ="& intInviterScore &"  where inviteAccount='" & updateMemID & "' and IsActive=0 and invitationCode=" & invitationCode
                    
                    conn.Execute(strSQLaddIsInvited)
                end if
			end if
			
			Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head><body bgcolor='#ffffff'>"
			Response.Write "<script language='javascript'>"
			Response.Write "alert(' Email 認證成功！您現在可以重新登入，開始使用入口網！');"
			Response.Write "location.replace('/mp.asp?mp=1');"
			Response.Write "</script>"
			Response.Write "</body></html>"
			Response.End()
		end if

	End If

	'認證失敗
	Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head><body bgcolor='#ffffff'>"
	Response.Write "<script language='javascript'>"
	Response.Write "alert(' Email 認證失敗！請重新申請或尋求協助！\n" & email & "');"
	Response.Write "this.close();"
	Response.Write "</script>"
	Response.Write "</body></html>"
	Response.End()
%>
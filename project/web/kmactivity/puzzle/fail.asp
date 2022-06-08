<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	Function GetSecureVal(param)
		If IsEmpty(param) Or param = "" Then
			GetSecureVal = param
			Exit Function
		End If
	
		If IsNumeric(param) Then
			GetSecureVal = CLng(param)
		Else
			GetSecureVal = Replace(CStr(param), "'", "''")
		End If
	End Function
	
	picNum = GetSecureVal(request("picNum"))
	loginid = GetSecureVal(request("loginid"))
	token = GetSecureVal(request("token"))
	p1 = GetSecureVal(request("p1"))
	p2 = GetSecureVal(request("p2"))
	p3 = GetSecureVal(request("p3"))
	p4 = GetSecureVal(request("p4"))
	p5 = GetSecureVal(request("p5"))
	p6 = GetSecureVal(request("p6"))
	p7 = GetSecureVal(request("p7"))
	p8 = GetSecureVal(request("p8"))
	p9 = GetSecureVal(request("p9"))
	p10 = GetSecureVal(request("p10"))
	p11 = GetSecureVal(request("p11"))
	p12 = GetSecureVal(request("p12"))
	p13 = GetSecureVal(request("p13"))
	p14 = GetSecureVal(request("p14"))
	p15 = GetSecureVal(request("p15"))
	p16 = GetSecureVal(request("p16"))
	p17 = GetSecureVal(request("p17"))
	p18 = GetSecureVal(request("p18"))
	p19 = GetSecureVal(request("p19"))
	p20 = GetSecureVal(request("p20"))
	p21 = GetSecureVal(request("p21"))
	p22 = GetSecureVal(request("p22"))
	p23 = GetSecureVal(request("p23"))
	p24 = GetSecureVal(request("p24"))
	p25 = GetSecureVal(request("p25"))
	p26 = GetSecureVal(request("p26"))
	p27 = GetSecureVal(request("p27"))
	p28 = GetSecureVal(request("p28"))
	p29 = GetSecureVal(request("p29"))
	p30 = GetSecureVal(request("p30"))
	p31 = GetSecureVal(request("p31"))
	p32 = GetSecureVal(request("p32"))
	p33 = GetSecureVal(request("p33"))
	p34 = GetSecureVal(request("p34"))
	p35 = GetSecureVal(request("p35"))
	p36 = GetSecureVal(request("p36"))
	diff = GetSecureVal(Request("difficult"))
	
	'檢查送過來的資料是否有齊全
	If picNum="" Or loginid="" Or token="" Or diff="" Then 
		Response.write "資料不同步，請按F5重新整理！"
		Response.End()
	End If
	for i = 1 to 36
		If eval("p"&i)="" Then 
			Response.write "資料不同步，請按F5重新整理！"
			Response.End()
		End If
	next
	
	'檢查送過來的資料是否有兩兩相同的狀況
	'For i = 1 to 36
	'	For j = i+1 to 36
	'		If cstr(eval("p"&i))=cstr(eval("p"&j)) Then 
	'			Response.write "資料不同步，請按F5重新整理！"
	'			Response.End()
	'		End If
	'	Next
	'Next
	
	set conn = server.createobject("adodb.connection")
	Conn.ConnectionString = application("ConnStrPuzzle")
	Conn.ConnectionTimeout=0
	Conn.CursorLocation = 3
	Conn.open
	
	'確認是否在登入時限內
	SQL = "Select * From sso Where Token='"& token &"' " 
	set rs = conn.execute (SQL)
	If DateAdd("n",30,RS("LastActiveTime"))<now() Then 
		Response.write "登入時間過久，請重新登入！"
		Response.End()
	End If
	
	'檢查遊戲狀態是否正確
	SQL = "Select * From GAMELOG Where LOGIN_ID='"& loginid &"' ORDER BY Ser_No DESC " 
	set rs = conn.execute (SQL)
	checkgame=0
	If Not RS.Eof Then 
		If cstr(RS("pic_id"))<>cstr(picNum) Then checkgame = checkgame + 1		
		If cstr(RS("LOGIN_ID"))<>cstr(loginid) Then checkgame = checkgame + 1
		If cstr(RS("pic1"))<>cstr(p1) Then checkgame = checkgame + 1
		If cstr(RS("pic2"))<>cstr(p2) Then checkgame = checkgame + 1
		If cstr(RS("pic3"))<>cstr(p3) Then checkgame = checkgame + 1
		If cstr(RS("pic4"))<>cstr(p4) Then checkgame = checkgame + 1
		If cstr(RS("pic5"))<>cstr(p5) Then checkgame = checkgame + 1
		If cstr(RS("pic6"))<>cstr(p6) Then checkgame = checkgame + 1
		If cstr(RS("pic7"))<>cstr(p7) Then checkgame = checkgame + 1
		If cstr(RS("pic8"))<>cstr(p8) Then checkgame = checkgame + 1
		If cstr(RS("pic9"))<>cstr(p9) Then checkgame = checkgame + 1
		If cstr(RS("pic10"))<>cstr(p10) Then checkgame = checkgame + 1
		If cstr(RS("pic11"))<>cstr(p11) Then checkgame = checkgame + 1
		If cstr(RS("pic12"))<>cstr(p12) Then checkgame = checkgame + 1
		If cstr(RS("pic13"))<>cstr(p13) Then checkgame = checkgame + 1
		If cstr(RS("pic14"))<>cstr(p14) Then checkgame = checkgame + 1
		If cstr(RS("pic15"))<>cstr(p15) Then checkgame = checkgame + 1
		If cstr(RS("pic16"))<>cstr(p16) Then checkgame = checkgame + 1
		If cstr(RS("pic17"))<>cstr(p17) Then checkgame = checkgame + 1
		If cstr(RS("pic18"))<>cstr(p18) Then checkgame = checkgame + 1
		If cstr(RS("pic19"))<>cstr(p19) Then checkgame = checkgame + 1
		If cstr(RS("pic20"))<>cstr(p20) Then checkgame = checkgame + 1
		If cstr(RS("pic21"))<>cstr(p21) Then checkgame = checkgame + 1
		If cstr(RS("pic22"))<>cstr(p22) Then checkgame = checkgame + 1
		If cstr(RS("pic23"))<>cstr(p23) Then checkgame = checkgame + 1
		If cstr(RS("pic24"))<>cstr(p24) Then checkgame = checkgame + 1
		If cstr(RS("pic25"))<>cstr(p25) Then checkgame = checkgame + 1
		If cstr(RS("pic26"))<>cstr(p26) Then checkgame = checkgame + 1
		If cstr(RS("pic27"))<>cstr(p27) Then checkgame = checkgame + 1
		If cstr(RS("pic28"))<>cstr(p28) Then checkgame = checkgame + 1
		If cstr(RS("pic29"))<>cstr(p29) Then checkgame = checkgame + 1
		If cstr(RS("pic30"))<>cstr(p30) Then checkgame = checkgame + 1
		If cstr(RS("pic31"))<>cstr(p31) Then checkgame = checkgame + 1
		If cstr(RS("pic32"))<>cstr(p32) Then checkgame = checkgame + 1
		If cstr(RS("pic33"))<>cstr(p33) Then checkgame = checkgame + 1
		If cstr(RS("pic34"))<>cstr(p34) Then checkgame = checkgame + 1
		If cstr(RS("pic35"))<>cstr(p35) Then checkgame = checkgame + 1
		If cstr(RS("pic36"))<>cstr(p36) Then checkgame = checkgame + 1		
	Else
		Response.write "資料不同步，請按F5重新整理！"
		Response.End()	
	End If
	
	If checkgame>2 Then 
		Response.write "資料不同步，請按F5重新整理！"
		Response.End()
	End If
	
	'寫入本次遊戲狀態		
	SQL="GAMELOG"
	Set RS = Server.CreateObject("ADODB.RecordSet")
  	RS.Open SQL,Conn,1,3
  	RS.Addnew
  	RS("pic_id")=picNum
	RS("LOGIN_ID")=loginid
	RS("pic1")=p1
	RS("pic2")=p2
	RS("pic3")=p3
	RS("pic4")=p4
	RS("pic5")=p5
	RS("pic6")=p6
	RS("pic7")=p7
	RS("pic8")=p8
	RS("pic9")=p9
	RS("pic10")=p10
	RS("pic11")=p11
	RS("pic12")=p12
	RS("pic13")=p13
	RS("pic14")=p14
	RS("pic15")=p15
	RS("pic16")=p16
	RS("pic17")=p17
	RS("pic18")=p18
	RS("pic19")=p19
	RS("pic20")=p20
	RS("pic21")=p21
	RS("pic22")=p22
	RS("pic23")=p23
	RS("pic24")=p24
	RS("pic25")=p25
	RS("pic26")=p26
	RS("pic27")=p27
	RS("pic28")=p28
	RS("pic29")=p29
	RS("pic30")=p30
	RS("pic31")=p31
	RS("pic32")=p32
	RS("pic33")=p33
	RS("pic34")=p34
	RS("pic35")=p35
	RS("pic36")=p36
	check=0
	for i = 1 to 36
		if cint(eval("p"&i))<>cint(i) then check = check + 1
	next
	RS("picState")="F"	
	RS("gametime")=now
	RS("difficult")=diff
	check="F"
  	RS.Update
  	RS.Close	
	
	'更新登入時間
	SQL = "update sso set LastActiveTime=getdate() Where Token='"& token &"' "	
	set rs = conn.execute (SQL)
	
	'本拼圖完成後把遊戲狀態清掉並將資料寫入game history中
	if check="F" Then 		
		SQL = "declare @LOGIN_ID varchar(50)"
        SQL = SQL & vbcrlf & "set @LOGIN_ID = '"& loginid &"'" 
        SQL = SQL & vbcrlf & "begin tran"
        SQL = SQL & vbcrlf & "	insert into GAMEHistory (pic_id, LOGIN_ID, pic1, pic2, pic3, pic4, pic5, pic6, pic7, pic8, pic9, pic10, pic11, pic12, pic13, pic14, pic15, pic16, pic17, pic18, pic19, pic20, pic21, "
        SQL = SQL & vbcrlf & "	  pic22, pic23, pic24, pic25, pic26, pic27, pic28, pic29, pic30, pic31, pic32, pic33, pic34, pic35, pic36, picState, gametime, difficult)"
        SQL = SQL & vbcrlf & ""
        SQL = SQL & vbcrlf & "	select pic_id, LOGIN_ID, pic1, pic2, pic3, pic4, pic5, pic6, pic7, pic8, pic9, pic10, pic11, pic12, pic13, pic14, pic15, pic16, pic17, pic18, pic19, pic20, pic21, "
        SQL = SQL & vbcrlf & "	  pic22, pic23, pic24, pic25, pic26, pic27, pic28, pic29, pic30, pic31, pic32, pic33, pic34, pic35, pic36, picState, gametime, difficult"
        SQL = SQL & vbcrlf & "	from GAMELOG where login_id = @LOGIN_ID"
        SQL = SQL & vbcrlf & ""
        SQL = SQL & vbcrlf & "	delete from GAMELOG where login_id = @LOGIN_ID"
        SQL = SQL & vbcrlf & "commit		"
		set rs1 = conn.execute (SQL)
		
		response.write "放棄成功！" '回給flash放棄成功	
	end if
	
	'rs.close()
	'rs1.close()
	conn.close()
%>
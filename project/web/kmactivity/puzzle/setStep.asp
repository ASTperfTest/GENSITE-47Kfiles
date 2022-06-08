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
		If RS("pic_id")<>picNum Then checkgame = checkgame + 1
		If RS("LOGIN_ID")<>loginid Then checkgame = checkgame + 1
		If RS("pic1")<>p1 Then checkgame = checkgame + 1
		If RS("pic2")<>p2 Then checkgame = checkgame + 1
		If RS("pic3")<>p3 Then checkgame = checkgame + 1
		If RS("pic4")<>p4 Then checkgame = checkgame + 1
		If RS("pic5")<>p5 Then checkgame = checkgame + 1
		If RS("pic6")<>p6 Then checkgame = checkgame + 1
		If RS("pic7")<>p7 Then checkgame = checkgame + 1
		If RS("pic8")<>p8 Then checkgame = checkgame + 1
		If RS("pic9")<>p9 Then checkgame = checkgame + 1
		If RS("pic10")<>p10 Then checkgame = checkgame + 1
		If RS("pic11")<>p11 Then checkgame = checkgame + 1
		If RS("pic12")<>p12 Then checkgame = checkgame + 1
		If RS("pic13")<>p13 Then checkgame = checkgame + 1
		If RS("pic14")<>p14 Then checkgame = checkgame + 1
		If RS("pic15")<>p15 Then checkgame = checkgame + 1
		If RS("pic16")<>p16 Then checkgame = checkgame + 1
		If RS("pic17")<>p17 Then checkgame = checkgame + 1
		If RS("pic18")<>p18 Then checkgame = checkgame + 1
		If RS("pic19")<>p19 Then checkgame = checkgame + 1
		If RS("pic20")<>p20 Then checkgame = checkgame + 1
		If RS("pic21")<>p21 Then checkgame = checkgame + 1
		If RS("pic22")<>p22 Then checkgame = checkgame + 1
		If RS("pic23")<>p23 Then checkgame = checkgame + 1
		If RS("pic24")<>p24 Then checkgame = checkgame + 1
		If RS("pic25")<>p25 Then checkgame = checkgame + 1
		If RS("pic26")<>p26 Then checkgame = checkgame + 1
		If RS("pic27")<>p27 Then checkgame = checkgame + 1
		If RS("pic28")<>p28 Then checkgame = checkgame + 1
		If RS("pic29")<>p29 Then checkgame = checkgame + 1
		If RS("pic30")<>p30 Then checkgame = checkgame + 1
		If RS("pic31")<>p31 Then checkgame = checkgame + 1
		If RS("pic32")<>p32 Then checkgame = checkgame + 1
		If RS("pic33")<>p33 Then checkgame = checkgame + 1
		If RS("pic34")<>p34 Then checkgame = checkgame + 1
		If RS("pic35")<>p35 Then checkgame = checkgame + 1
		If RS("pic36")<>p36 Then checkgame = checkgame + 1
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
	if check = 0 Then 
		RS("picState")="Y"
		check = "Y"
	else
		RS("picState")="N"
	end if	
	RS("gametime")=now
	RS("difficult")=diff
  	RS.Update
  	RS.Close	
	
	'更新登入時間
	SQL = "update sso set LastActiveTime=getdate() Where Token='"& token &"' "	
	set rs = conn.execute (SQL)
	
	SQL = "Select * From ACCOUNT Where Login_ID = '"& loginid &"' "	
	set rs = conn.execute (SQL)
	
	'扣掉點數
	If Not RS.Eof Then 
		energy = RS("Energy")
	'	energy = energy - 1
	'	SQL = "update ACCOUNT set Energy= '"& energy &"' Where Login_ID = '"& loginid &"' "	
	'	set rs1 = conn.execute (SQL)		
	End If
	
	'本拼圖完成後把遊戲狀態清掉並將資料寫入game history中
	if check="Y" Then 		
		SQL = "Select * From GAMELOG Where LOGIN_ID='"& loginid &"'" 
		set rs = conn.execute (SQL)
		While Not RS.Eof 
			SQL="GAMEHistory"
			Set RS1 = Server.CreateObject("ADODB.RecordSet")
  			RS1.Open SQL,Conn,1,3
  			RS1.Addnew
  			RS1("pic_id")=RS("pic_id")
			RS1("LOGIN_ID")=RS("LOGIN_ID")
			RS1("pic1")=RS("pic1")
			RS1("pic2")=RS("pic2")
			RS1("pic3")=RS("pic3")
			RS1("pic4")=RS("pic4")
			RS1("pic5")=RS("pic5")
			RS1("pic6")=RS("pic6")
			RS1("pic7")=RS("pic7")
			RS1("pic8")=RS("pic8")
			RS1("pic9")=RS("pic9")
			RS1("pic10")=RS("pic10")
			RS1("pic11")=RS("pic11")
			RS1("pic12")=RS("pic12")
			RS1("pic13")=RS("pic13")
			RS1("pic14")=RS("pic14")
			RS1("pic15")=RS("pic15")
			RS1("pic16")=RS("pic16")
			RS1("pic17")=RS("pic17")
			RS1("pic18")=RS("pic18")
			RS1("pic19")=RS("pic19")
			RS1("pic20")=RS("pic20")
			RS1("pic21")=RS("pic21")
			RS1("pic22")=RS("pic22")
			RS1("pic23")=RS("pic23")
			RS1("pic24")=RS("pic24")
			RS1("pic25")=RS("pic25")
			RS1("pic26")=RS("pic26")
			RS1("pic27")=RS("pic27")
			RS1("pic28")=RS("pic28")
			RS1("pic29")=RS("pic29")
			RS1("pic30")=RS("pic30")
			RS1("pic31")=RS("pic31")
			RS1("pic32")=RS("pic32")
			RS1("pic33")=RS("pic33")
			RS1("pic34")=RS("pic34")
			RS1("pic35")=RS("pic35")
			RS1("pic36")=RS("pic36")
			RS1("picState")=RS("picState")		
			RS1("gametime")=RS("gametime")
		  	RS1.Update
		  	RS1.Close
			
			RS.Movenext
		Wend
		'把遊戲狀態清掉
		SQL = "Delete From GAMELOG Where LOGIN_ID='"& loginid &"'" 
		set rs1 = conn.execute (SQL)
		response.write "0" '回給flash拼圖完成
	elseif energy = 0 Then 
		response.write "x" '回給flash點數不足
	else
		response.write energy '回給flash剩餘點數
	end if
	
	rs.close()
	'rs1.close()
	conn.close()
%>
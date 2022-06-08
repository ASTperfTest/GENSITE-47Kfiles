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
	
	If GetSecureVal(Request("token"))="" Then 
		Response.write "請先登入會員！"		
		Response.End()
	End If
	
	set conn = server.createobject("adodb.connection")
	Conn.ConnectionString = application("ConnStrPuzzle")
	Conn.ConnectionTimeout=0
	Conn.CursorLocation = 3
	Conn.open
	
	'取得登入時間
	SQL = "Select * From sso Where Token='"& GetSecureVal(Request("token")) &"' "		
	set rs = conn.execute (SQL)
	
	If Not RS.Eof Then 
		If DateAdd("n",30,RS("LastActiveTime"))<now() Then 
			Response.write "登入時間過久，請重新登入！"
		Else			
			SQL = "Select * From GAMELOG Where Login_ID = '"& RS("Login_ID") &"' ORDER BY Ser_No DESC  "		
			set rs1 = conn.execute (SQL)
			If Not RS1.Eof Then
				Response.write RS1("pic1") & "<->" & RS1("pic2") & "<->" & RS1("pic3") & "<->" & RS1("pic4") & "<->" & RS1("pic5") & "<->" & RS1("pic6") & "<->" & RS1("pic7") & "<->" & RS1("pic8") & "<->" & RS1("pic9") & "<->" & RS1("pic10") & "<->" & RS1("pic11") & "<->" & RS1("pic12") & "<->" & RS1("pic13") & "<->" & RS1("pic14") & "<->" & RS1("pic15") & "<->" & RS1("pic16") & "<->" & RS1("pic17") & "<->" & RS1("pic18") & "<->" & RS1("pic19") & "<->" & RS1("pic20") & "<->" & RS1("pic21") & "<->" & RS1("pic22") & "<->" & RS1("pic23") & "<->" & RS1("pic24") & "<->" & RS1("pic25") & "<->" & RS1("pic26") & "<->" & RS1("pic27") & "<->" & RS1("pic28") & "<->" & RS1("pic29") & "<->" & RS1("pic30") & "<->" & RS1("pic31") & "<->" & RS1("pic32") & "<->" & RS1("pic33") & "<->" & RS1("pic34") & "<->" & RS1("pic35") & "<->" & RS1("pic36") 
			Else
				Response.write "資料不同步，請按F5重新整理！"
			End If
		End If
	Else
		Response.write "您還未登入！"
	End If
	
	conn.close()
%>
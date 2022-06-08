<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%	
	set conn = server.createobject("adodb.connection")
	Conn.ConnectionString = application("ConnStrPuzzle")
	Conn.ConnectionTimeout=0
	Conn.CursorLocation = 3
	Conn.open

	SQL = "Select Count(*) AS Total From PICDATA"
	set RS = conn.execute (SQL)
	picNumbers=RS("Total")'宣告現有圖庫數量
	
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

	Function getpic()
		'randomize 		
		'x = int(rnd()*puzzleN)+1
		'x = Right("00" & x ,3) '亂數取圖片編號
		'x="001" '測試用拼圖編號
		SQL = "Select Top 1 * From PICDATA Where Pic_Open='Y' Order by NEWID()"  ' And (CONVERT(char(8),gametime,11)=CONVERT(char(8),GETDATE(),11))						
		set rs1 = conn.execute (SQL)
		getpic=RS1("Ser_No")
	End Function
	
	Function searchData(loginid,picno)		
		SQL = "Select * From GAMEHistory Where Login_ID = '"& loginid &"' And pic_id='"& picno &"' And (picState='Y' Or picState='F')"  ' And (CONVERT(char(8),gametime,11)=CONVERT(char(8),GETDATE(),11))						
		set rs1 = conn.execute (SQL)	
		If Not RS1.Eof Then 
			searchData=1'有玩過此張拼圖
		Else
			searchData=0
		End If
	End Function
	
	If GetSecureVal(Request("token"))="" Then 
		Response.write "請先登入會員！"		
		Response.End()
	End If	
	
	'取得登入時間
	SQL = "Select * From sso Where Token='"& GetSecureVal(Request("token")) &"' "		
	set RS = conn.execute (SQL)
	
	If Not RS.Eof Then 
		If DateAdd("n",30,RS("LastActiveTime"))<now() Then 
			Response.write "登入時間過久，請重新登入！"
		Else
			'先檢查是否己經玩完所有的圖片了
			SQL = "Select count(*) As Total From GAMEHistory Where Login_ID = '"& RS("Login_ID") &"' And (picState='Y' Or picState='F') "
			set rs1 = conn.execute (SQL)
			If Not RS1.eof Then 
				If cint(RS1("Total"))= picNumbers Then 
					Response.write "您今天己完成所有圖庫，請明日再登入遊戲。"
					Response.End()
				End If
			End If
			'先檢查是否今天己經玩超過圖庫總數
			SQL = "Select count(*) As Total From GAMEHistory Where Login_ID = '"& RS("Login_ID") &"' And (CONVERT(char(8),gametime,11)=CONVERT(char(8),GETDATE(),11)) And (picState='Y' Or picState='F') "			
			set rs1 = conn.execute (SQL)
			If Not RS1.eof Then 
				If cint(RS1("Total"))= picNumbers Then 
					Response.write "您今天己完成所有圖庫，請明日再登入遊戲。"
					Response.End()
				Elseif cint(RS1("Total")) >=5 Then
					Response.write "您今天己完成5個拼圖，請明日再登入遊戲。"
					Response.End()					
				Else
					picid=getpic() '取得隨機圖片編號
					checkpic=searchData(RS("Login_ID"),picid)'檢查是否有玩過此拼圖
				End If		
			End If			
			while checkpic=1
				picid=getpic()
				checkpic=searchData(RS("Login_ID"),picid)
			Wend
			SQL = "Select * From ACCOUNT Where Login_ID = '"& RS("Login_ID") &"'  "		
			set rs1 = conn.execute (SQL)
			SQL = "Select * From GAMELOG Where Login_ID = '"& RS("Login_ID") &"'  "		
			set rs2 = conn.execute (SQL)
			difficult=""
			If Not RS2.Eof Then 				
				Response.write "1<->" '有玩到一半拼圖的狀態
				picid = RS2("pic_id")
				difficult = RS2("difficult")
			Else	
				Response.write "0<->" '新的拼圖picid=PICDATA.Ser_No
			End If			
			Response.write picid & "<->"
			SQL = "Select * From PICDATA Where Ser_No = '"& picid &"'  "		
			set RS3 = conn.execute (SQL)
			If Not RS3.eof Then Response.write RS3("Pic_No") & "<->" & RS3("Pic_Name")
			Response.write "<->" & RS1("Login_ID") & "<->" 
			If RS1("NICKNAME")="" Then 
				If len(RS1("REALNAME"))=1 Then 
					Response.write RS1("REALNAME")
				Else
					Response.write left(RS1("REALNAME"),1)&"＊"&right(RS1("REALNAME"),len(RS1("REALNAME"))-2)
				End If
			Else
				Response.write RS1("NICKNAME")
			End If
			Response.write "<->" & RS1("Energy") 
			If Not RS3.eof Then Response.write "<->" & RS3("Pic_Link")
			If difficult<>"" Then Response.write "<->" & difficult
			SQL = "Select count(*) AS Total From GAMELOG Where pic_id = '"& picid &"' And LOGIN_ID='"& RS("Login_ID") &"'"
			set RS3 = conn.execute (SQL)
			Response.write "<->"& cint(RS3("Total"))-1
		End If
	Else
		Response.write "您還未登入！"
	End If
	rs.close
	'rs1.close
	'rs2.close
	'rs3.close
	conn.close()
%>
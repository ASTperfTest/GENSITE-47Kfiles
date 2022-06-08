<%
FUNCTION pkStr (s, endchar)
  if s="" then
	pkStr = "null" & endchar
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	pkStr="'" & s & "'" & endchar
  end if
END FUNCTION

  Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open session("ODBCDSN")
	ptime = request("ptime")
	
	tone  = "0"
	if request("tone") = "1" then
	 tone = request("toneselect")
	else
	 tone  = "0"
	end if
	
	name = pkStr(request("name"),"")
	
	
'	response.write pkStr(request("name"),"")
''  response.end
  'xSql = "INSERT INTO CuDTGeneric(iBaseDSD,iCtUnit,sTitle,xBody,xPostDate,xKeyword,iDept,iEditor,fctupublic,deditDate,xnewWindow,showType)" _
  '  	& " VALUES(38,651," & name & ",'',getdate(),'',0,'hyweb','Y',getdate(),'N',1)"
  xSql = "INSERT INTO CuDTGeneric(iBaseDSD,iCtUnit,sTitle,xBody,xPostDate,xKeyword,iDept,iEditor,fctupublic,deditDate,xnewWindow,showType)" _
    	& " VALUES(38,651," & name & ",'',getdate(),'',0,'hyweb','Y',rtrim(ltrim(str(year(getdate())))) + '/' + rtrim(ltrim(str(month(getdate())))) +  '/' + rtrim(ltrim(str(day(getdate())))),'N',1)"
    	
  sql = "set nocount on;"&xSql&"; select @@IDENTITY as NewID"    
		'debugPrint Sql
		'response.write 	fileName & "<BR>"
    'response.write sql	& "<BR>"
    'response.end
  set RSx = conn.Execute(sql)
  xNewIdentity = RSx(0)    	
  money = request("money")
  rate = money - ptime
  str="insert into flashGame(giCuItem,name,age,sex,ptimemin,ptimesec,ptime,email,hakkalangtone,money,rate) values(" & xNewIdentity & "," & name  & "," & pkStr(request("age"),"") & "," & pkStr(request("sex"),"") & "," & Cint((ptime-(ptime mod 60))/60) & "," & Cint(ptime mod 60) & "," & ptime & "," & pkStr(request("email"),"") & "," & pkStr(tone,"") & "," & money & "," & rate & ")"
   
  conn.execute(str)
  
  
  'SQL="select top 20 * from flashGame order by rate desc,giCuItem desc"
  SQL="select top 1 * from flashGame where name = " & name   & " and money > " & money & ""
  'response.write SQL
  'response.end
  SET gamedata=conn.execute(SQL)
  bfind = "N"
  while not gamedata.eof
    'response.write xNewIdentity
    'response.write gamedata("giCuItem")
    'if cstr(xNewIdentity) = cstr(gamedata("giCuItem")) then
      bfind = "Y"
    'end if
    gamedata.moveNext
  wend
  
  set conn = Nothing
  response.write "&IN=" & bfind
%>

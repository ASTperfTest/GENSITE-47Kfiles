<%
  Response.CacheControl = "no-cache" 
  Response.AddHeader "Pragma", "no-cache" 
  Response.Expires = -1  
  
  Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open session("ODBCDSN")
	SQL="select top 1 * from Member where account='" & request("userIdno") & "' and passwd='" & request("userEmail") & "'"
  SET gamedata=conn.execute(SQL)
  i = 0
  if not gamedata.eof then
	   count = gamedata("logincount")+1
	   
	   Today = Year(Now()) & "/" & Month(Now()) & "/" & Day(Now()) 
	   
	   'SQL = "select * from MemberFlashLoginLog where account='" & request("userIdno") & "' and LoginDatetime >= '" & Today & "' and LoginDatetime < DATEADD(day, 1, '" & Today & "')"

        SQL = "select top 1 * from flashGame,CuDTGeneric where CuDTGeneric.iCuItem=flashGame.giCuItem and name = '" & request("userIdno") & "' order by rate desc" 
        
        SET highMoneyrs=conn.execute(SQL)
        
        highMoney = 0
        if not highMoneyrs.eof then
            highMoney = highMoneyrs("money")
        end if

	    SQL = "select * from flashGame,CuDTGeneric where CuDTGeneric.iCuItem=flashGame.giCuItem and deditDate= (rtrim(ltrim(str(year(getdate())))) + '/' + rtrim(ltrim(str(month(getdate())))) +  '/' + rtrim(ltrim(str(day(getdate()))))) and name = '" & request("userIdno") & "' order by rate desc" 
	   
	   'Response.Write(SQL)
	   SET todaylogin=conn.execute(SQL)
	   
	   nCount = 0
	   
	    while not todaylogin.eof
	        nCount = nCount + 1
	        todaylogin.moveNext
        wend
	   'response.Write SQL & nCount
	   if nCount > 3 then
	     Response.Write("success=true&count=" & count & "&realname=" & gamedata("realname") & "&relogin=true&highMoney=" & highMoney & "&")
	   else	   
	     Response.Write("success=true&count=" & count & "&realname=" & gamedata("realname") & "&relogin=false&highMoney=" & highMoney & "&")
	     conn.execute("update Member set logincount=logincount+1 where  account='" & request("userIdno") & "' and passwd='" & request("userEmail") & "'")
       	     conn.execute("insert MemberFlashLoginLog(account) values('" & request("userIdno") & "')")
     	   end if	   
  else
    Response.Write("success=false&")
  end if
  
  set conn = Nothing
%>
  

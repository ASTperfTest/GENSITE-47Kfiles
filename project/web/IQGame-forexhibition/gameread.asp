<%
  Response.CacheControl = "no-cache" 
  Response.AddHeader "Pragma", "no-cache" 
  Response.Expires = -1
  
  
  Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open session("ODBCDSN")
	SQL="select top 20 * from flashGame order by rate desc,giCuItem desc"
  SET gamedata=conn.execute(SQL)
  i = 0
  while not gamedata.eof
    i = i + 1
    Response.Write("&name" & i & "=" & gamedata("name"))
    Response.Write("&age" & i & "=" & gamedata("age"))
    Response.Write("&sex" & i & "=" & gamedata("sex"))
    Response.Write("&ptimemin" & i & "=" & gamedata("ptimemin"))
    Response.Write("&ptimesec" & i & "=" & gamedata("ptimesec"))
    Response.Write("&rate" & i & "=" & gamedata("rate"))
    gamedata.moveNext
  wend
  
  set conn = Nothing
%>

'			response.write sql & sqlValue & ")<HR>"
			conn.execute sql & sqlValue & ")"
		next
'		response.end
		response.redirect session("pickReturnAPURI")
	end if

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 fSql=Request.QueryString("strSql")

if fSql="" then

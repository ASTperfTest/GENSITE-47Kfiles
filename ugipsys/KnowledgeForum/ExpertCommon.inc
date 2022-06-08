<%
	Function GetsTitle( id, atype )
		Dim str : str = ""
		sql = "SELECT sTitle FROM CuDTGeneric WHERE icuitem = " & id
		set rs = conn.execute(sql)
		if not rs.eof then
			str = rs("sTitle")
		end if
		rs.close
		set rs = nothing
		if atype = "1" then str = "專家補充-" & id & "-" & str
		GetsTitle = str
	End Function
	
	Function GetRand()		
		Dim yea : yea = year(now)
		Dim mon : mon = month(now)		
		Dim da : da = day(now)		
		Dim hou : hou = hour(now)
		Dim min : min = minute(now)
		Dim sec : sec = second(now)
		Randomize
		temp = rnd * 100 mod 6
		if temp = 0 then
			GetRand = yea & mon & da & hou & min & sec 
		elseif temp = 1 then
			GetRand = mon & da & sec & min & hou & yea 
		elseif temp = 2 then
			GetRand = da & min & yea & sec & mon & hou 
		elseif temp = 3 then
			GetRand = hou & mon & min & da & yea & sec 
		elseif temp = 4 then
			GetRand = min & sec & da & hou & yea & mon 
		elseif temp = 5 then
			GetRand = sec & hou & da & mon & min & yea 
		end if
	End Function
	
	Function GetEmailBodyFromFile()
		CONST ForReading = 1, ASCII = 0
		Dim objFSO, objTextStream
		FileName = "MailBody.txt"
		FilePath = Server.MapPath(FileName)
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		if objFSO.FileExists(FilePath) then
			Set objTextStream = objFSO.OpenTextFile(FilePath, ForReading, False, ASCII)
		else
			response.write "mailbody.txt does not exist"
			response.end
		end if
		AllStr = objTextStream.ReadAll
		objTextStream.close
		set objTextStream = nothing
		set objFSO = nothing
		GetEmailBodyFromFile = AllStr
	End Function
	
	Function ReplaceTag( body, mailbody, expertId, questionId, myrand, icuitem )
		Dim realname, email
		sql = "SELECT realname, email FROM Member WHERE account = '" & expertId & "'"
		set rs = conn.execute(sql)
		if not rs.eof then
			realname = rs("realname")
			email = rs("email")
		end if
		rs.close
		set rs = nothing
		body = replace(body, "{receiver}", realname & "&nbsp;" & email)
		body = replace(body, "{expertname}", realname)
		body = replace(body, "{mailbody}", mailbody)
		Dim qtitle, qbody, qlink
		sql = "SELECT sTitle, xBody, topCat FROM CuDTGeneric WHERE icuitem = " & questionId
		set rs = conn.execute(sql)
		if not rs.eof then
			qtitle = rs("sTitle")
			qbody = rs("xBody")
			qlink = session("myWWWSiteURL") & "/knowledge/knowledge_cp.aspx?ArticleId=" & questionId & "&ArticleType=" & rs("topCat")
		end if 
		rs.close
		set rs = nothing		
		body = replace(body, "{questiontitle}", qtitle)
		body = replace(body, "{questioncontent}", qbody)
		body = replace(body, "{questionlink}", qlink)
		body = replace(body, "{replylink}", session("myWWWSiteURL") & "/Knowledge/KnowledgeExpertReply.aspx?ArticleId=" & questionId & "&DArticleId=" & icuitem & "&rand=" & myrand & "&expertId=" & expertId)		
		ReplaceTag = body
	End Function
	
	Function GetExpertEmail( expertId )
		dim email
		sql = "SELECT email FROM Member WHERE account = '" & expertId & "'"
		set rs = conn.execute(sql)
		if not rs.eof then
			email = rs("email")
		end if 
		rs.close
		set rs = nothing
		GetExpertEmail = email
	End Function

%>
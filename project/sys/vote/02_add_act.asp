<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	'HTProgCap = "功能管理"
	'HTProgFunc = "問卷調查"
	HTProgCode = "GW1_vote01"
	'HTProgPrefix = "mSession"
	response.expires = 0
%>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<%	
	
	subject = replace(trim(request("subject")), "'", "''")
	if subject = "" then
		response.write "<body onload=JavaScript:alert('請輸入主題名稱！');history.go(-1);>"
		response.end		
	end if
	
	
	bdate = request("bdate_y") & "/" & request("bdate_m") & "/" & request("bdate_d")
	edate = request("edate_y") & "/" & request("edate_m") & "/" & request("edate_d")	
	if not ( isDate(bdate) and isDate(edate) )  then
		response.write "<body onload=javascript:alert('起訖日期輸入不正確！');history.go(-1);>"
		response.end
	end if
	
	
	submit_str = trim(request("submit"))
	subjectid = request("subjectid")
	

	if submit_str = "刪除本問卷" and subjectid <> "" then
		
		sql = "" & _
			" delete from m011 where m011_subjectid = " & subjectid & "; " & _
			" delete from m012 where m012_subjectid = " & subjectid & "; " & _
			" delete from m013 where m013_subjectid = " & subjectid & "; " & _
			" delete from m014 where m014_subjectid = " & subjectid & "; " & _
			" delete from m015 where m015_subjectid = " & subjectid & "; " & _
			" delete from m016 where m016_subjectid = " & subjectid
		conn.execute(sql)	
	
		response.write "<script language='javascript'>alert('刪除完畢！');location.replace('02.asp');</script>"
		response.end
	end if
	
	
	notetype_str = trim(request("notetype"))
	notetype = ""
	for i = 1 to 9
		if instr(notetype_str, i) > 0 then
			notetype = notetype & "1"
		else
			notetype = notetype & "0"
		end if
	next
	
	
	if (submit_str = "繼續修改問卷內容" or submit_str = "修改回列表") and subjectid <> "" then
		sql = "" & _
			" update m011 set " & _
			" m011_subject = '" & subject & "', " & _
			" m011_questno = " & request("questno") & ", " & _
			" m011_bdate = '" & bdate & "', " & _
			" m011_edate = '" & edate & "', " & _
			" m011_online = '" & request("online") & "', " & _
			" m011_notetype = '" & notetype & "', " & _
			" m011_questionnote = '" & replace(trim(request("questionnote")), "'", "''") & "', " & _
			" m011_haveprize = '" & request("haveprize") & "', " & _
			" m011_jumpquestion = '" & request("jumpquestion") & "', " & _
			" m011_onlyonce = '" & request("onlyonce") & "', " & _
			" m011_modifyuser = '" & session("NetUser") & "', " & _
			" m011_updatetime = getdate(), " & _
			" m011_km_online = '" & request("kmonline") & "' " & _
			" where " & _
			" m011_subjectid = " & subjectid
		conn.execute(sql)
		
		if submit_str = "繼續修改問卷內容" then
			response.redirect "02_add1.asp?subjectid=" & subjectid
		elseif submit_str = "修改回列表" then
			response.write "<script language='javascript'>alert('內容已修改');location.replace('02.asp');</script>"
		end if
		response.end
	end if
	
	
	set rs = conn.execute("select IsNull(Max(m011_subjectid), 0) from m011")
	subjectid = rs(0) + 1	
	
	sql = "" & _
		" insert into m011 ( " & _
		" m011_subjectid, " & _
		" m011_subject, " & _
		" m011_questno, " & _
		" m011_bdate, " & _
		" m011_edate, " & _
		" m011_online, " & _
		" m011_notetype, " & _
		" m011_questionnote, " & _
		" m011_haveprize, " & _
		" m011_jumpquestion, " & _
		" m011_onlyonce, " & _
		" m011_createuser, " & _
		" m011_createtime, " & _
		" m011_modifyuser, " & _
		" m011_updatetime, " & _
		" m011_km_online " & _
		" ) values ( " & _
		subjectid & ", " & _
		" '" & subject & "', " & _
		request("questno") & ", " & _
		" '" & bdate & "', " & _
		" '" & edate & "', " & _
		" '" & request("online") & "', " & _
		" '" & notetype & "', " & _
		" '" & Replace(trim(request("questionnote")), "'", "''") & "', " & _
		" '" & request("haveprize") & "', " & _
		" '" & request("jumpquestion") & "', " & _
		" '" & request("onlyonce") & "', " & _
		" '" & session("NetUser") & "', " & _
		" getdate(), " & _
		" '" & session("NetUser") & "', "& _
		" getdate(), " & _
		" '" & request("kmonline") & "' "& _
		" ) "
	conn.execute(sql)
	
	
	response.redirect "02_add1.asp?subjectid=" & subjectid
%>
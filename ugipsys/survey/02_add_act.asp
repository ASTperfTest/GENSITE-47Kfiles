<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
	subjectid = request("subjectid")
'     createuser="mike"
        subject = Replace(trim(request("subject")), "'", "''")
	if subject = "" then
		response.write "<body onload=JavaScript:alert('請輸入主題名稱！');history.go(-1);>"
		response.end		
	end if
	set bd = conn.execute("select m011_bdate from m011 where m011_subjectid = " & subjectid)
	set ed = conn.execute("select m011_edate from m011 where m011_subjectid = " & subjectid)
	bdate = bd(0)
	edate = ed(0)

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
			" delete from m014 where m014_subjectid = " & subjectid
		conn.execute(sql)	
		
	
		response.write "<script language='javascript'>alert('刪除完畢！');location.replace('adm_inves.asp');</script>"
		response.end
	end if
	
	
	notetype_str = trim(request("notetype"))
	notetype = ""
'	response.write trim(request("notetype"))& "<HR>"
	notetype_str = " " & notetype_str & ","
	haveprize = request("haveprize")
	if haveprize = "1" then
'		response.write notetype_str & "<HR>"
		for i = 1 to 4
			if Instr(notetype_str, " " & i& ",") > 0 then
				notetype = notetype & "1"
			else
				notetype = notetype & "0"
			end if

		next

	end if
'	response.write notetype
'	response.end
	
	if submit_str = "繼續修改問卷內容" and subjectid <> "" then
		sql = "" & _
			" update m011 set " & _
			" m011_subject = '" & subject & "', " & _
			" m011_questno = " & request("questno") & ", " & _
			" m011_notetype = '" & notetype & "', " & _
			" m011_haveprize = '" & haveprize & "', " & _
			" m011_onlyonce = '" & request("onlyonce") & "'" & _
			" where " & _
			" m011_subjectid = " & subjectid
		conn.execute(sql)
	
		response.redirect "02_add1.asp?subjectid=" & subjectid
		response.end
	end if
	
	
	set rs = conn.execute("select IsNull(Max(m011_subjectid), 0) from m011")
	subjectid = rs(0) + 1	
	
	sql = "" & _
		" insert into m011 ( " & _
		" m011_subjectid, " & _
		" m011_subject, " & _
		" m011_questno, " & _
		" m011_notetype, " & _
		" m011_questionnote, " & _
		" m011_haveprize, " & _
		" m011_onlyonce, " & _
		" m011_hide, " & _
		" m011_createuser, " & _
		" m011_createtime, " & _
		" m011_modifyuser, " & _
		" m011_updatetime " & _
		" ) values ( " & _
		subjectid & ", " & _
		" '" & subject & "', " & _
		request("questno") & ", " & _
		" '" & notetype & "', " & _
		" '" & Replace(trim(request("questionnote")), "'", "''") & "', " & _
		" '" & haveprize & "', " & _
		" '" & request("onlyonce") & "', " & _
		" '" & request("askhide") & "', " & _
		" '" & createuser & "', " & _
		" getdate(), " & _
		" '" & createuser & "', "& _
		" getdate() " & _
		" ) "

	conn.execute(sql)
	
	
	response.redirect "02_add1.asp?subjectid=" & subjectid
%>

<%
'---
Dim NowTime
NowTime = Now
NowTime = CInt(hour(NowTime))
if NowTime >= 6 And NowTime < 12 Then
	response.write "<NowTime>AM</NowTime>"
elseif NowTime >= 12 And NowTime < 18 Then
	response.write "<NowTime>PM</NowTime>"
else
	response.write "<NowTime>NT</NowTime>"
end if

'---current url---
redircetUrl = left(request.querystring, instr(request.querystring, "memID") - 2)
redircetUrl = replace(redircetUrl, "&", ";")
if instr(Request.ServerVariables("URL"), "xdcp2.asp") > 0 then
	response.write "<redirectUrl><![CDATA[/ct.asp?" & redircetUrl & "]]></redirectUrl>"
elseif instr(Request.ServerVariables("URL"), "xdlp.asp") > 0 then
	response.write "<redirectUrl><![CDATA[/lp.asp?" & redircetUrl & "]]></redirectUrl>"
elseif instr(Request.ServerVariables("URL"), "xdmp.asp") > 0 then
	response.write "<redirectUrl><![CDATA[/mp.asp?" & redircetUrl & "]]></redirectUrl>"
end if
'---

'---activity---
response.write "<activity>"
'---Is activity in the time range---
sql = "SELECT * FROM Activity WHERE (GETDATE() BETWEEN ActivityStartTime AND ActivityEndTime) AND ActivityId = '" & session("ActivityId") & "'"
Set actrs = conn.Execute(sql)
If Not actrs.Eof Then
	response.write "<active>true</active>"
Else
	response.write "<active>false</active>"
End If
actrs.close
set actrs = nothing

'---get the user activity grade---
if request("memID") <> "" Then	
	sql = "SELECT (DiscussCheckGrade + QuestionCheckGrade) AS Grade FROM ActivityMemberNew WHERE MemberId = '" & request("memID") & "'"
	Set actmemrs = conn.Execute(sql)
	If Not actmemrs.Eof Then
		response.write "<ActivityGrade>" & actmemrs("Grade") & "</ActivityGrade>"
	Else
		response.write "<ActivityGrade>0</ActivityGrade>"
	End If
	actmemrs.close
	Set actmemrs = Nothing
End If
response.write "</activity>"

	response.write "<webActivity>"

	sql = "SELECT * FROM m011 WHERE GETDATE() BETWEEN m011_bdate AND m011_edate AND m011_online = '1'"
	set rs = conn.execute(sql)
	If rs.eof Then 
		response.write "<active>false</active>"
		response.write "<id></id>"
	else
		response.write "<active>true</active>"
		response.write "<id>" & rs ("m011_subjectid") & "</id>"
	end if
	rs.close
	set rs = nothing

	response.write "</webActivity>"
	
%>	
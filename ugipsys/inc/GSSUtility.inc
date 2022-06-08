<%
function GetLoginInfo(memID, gstyle)
	response.write "<login>"
	Dim memName, memnickname, memlogincount
	if memID <> "" Then		
		sql = "SELECT realname, ISNULL(nickname, '') AS nickname, ISNULL(logincount, 0) AS logincount " & _
					"FROM Member where account = '" & memID & "'"
		Set memrs = conn.Execute(sql)
		If Not memrs.EOF Then
			response.write "<status>true</status>"	
			memName = trim(memrs("realname"))
			memnickname = trim(memrs("nickname"))								
			memlogincount = memrs("logincount")
		else 
				response.write "<status>false</status>"	
		End If
	else
		response.write "<status>false</status>"
	End If
	response.write "<memID>" & memID & "</memID>"
	response.write "<memName><![CDATA[" & memName & "]]></memName>"
	response.write "<memNickName><![CDATA[" & memnickname & "]]></memNickName>"
	response.write "<memLoginCount>" & memlogincount & "</memLoginCount>"		
	response.write "<gstyle>" & gstyle & "</gstyle>"				
	response.write "</login>"
end function
%>

<%
	function setInvitationCodeURL()
		inviteCodeURL = ""
		if request("Fcode") <> ""  then
			inviteCodeURL = "&Fcode=" & request("Fcode")
		end if
		if request("Ecode") <> ""  then
			inviteCodeURL = "&Ecode=" & request("Ecode")
		end if
		if request("Ucode") <> ""  then
			inviteCodeURL = "&Ucode=" & request("Ucode")
		end if
		'response.write(inviteCodeURL)
        setInvitationCodeURL = inviteCodeURL
	end function

	function getInvitationCode()
		inviteCode = ""
		if request("Fcode") <> ""  then
			inviteCode = request("Fcode")
		end if
		if request("Ecode") <> ""  then
			inviteCode = request("Ecode")
		end if
		if request("Ucode") <> ""  then
			inviteCode = request("Ucode")
		end if
        getInvitationCode = inviteCode
	end function

	function getInvitationCodeType()
		inviteCodeType = ""
		if request("Fcode") <> ""  then
			inviteCodeType = 1
		end if
		if request("Ecode") <> ""  then
			inviteCodeType = 2
		end if
		if request("Ucode") <> ""  then
			inviteCodeType = 0
		end if
        getInvitationCodeType = inviteCodeType
	end function
%>
<%
sub xpCondition
	if request.form("htx_psnID") <> "" then
		whereCondition = replace("htx.psnID LIKE '%{0}%'", "{0}", request.form("htx_psnID") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_pName") <> "" then
		whereCondition = replace("htx.pName LIKE '%{0}%'", "{0}", request.form("htx_pName") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_birthDayS") <> "" then
		rangeS = request.form("htx_birthDayS")
		rangeE = request.form("htx_birthDayE")
		if rangeE = "" then	rangeE=rangeS
		whereCondition = replace("htx.birthDay BETWEEN '{0}' AND '{1}'", "{0}", rangeS)
		whereCondition = replace(whereCondition, "{1}", rangeE)
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_cDateS") <> "" then
		rangeS = request.form("htx_cDateS")
		rangeE = request.form("htx_cDateE")
		if rangeE = "" then	rangeE=rangeS
		whereCondition = replace("htx.cDate BETWEEN '{0}' AND '{1}'", "{0}", rangeS)
		whereCondition = replace(whereCondition, "{1}", rangeE)
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_myOrg") <> "" then
		whereCondition = replace("htx.myOrg LIKE '%{0}%'", "{0}", request.form("htx_myOrg") )
		fSql = fSql & " AND " & whereCondition
	end if
end sub
%>

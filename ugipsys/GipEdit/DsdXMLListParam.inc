<%
sub xpCondition
	if request.form("htx_sTitle") <> "" then
		whereCondition = replace("ghtx.stitle LIKE {0}", "{0}", _
				pkStr("%"&request.form("htx_sTitle")&"%",""))
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_iDept") <> "" then
		whereCondition = replace("ghtx.iDept = {0}", "{0}", pkstr(request.form("htx_iDept"),"") )
		fSql = fSql & " AND " & whereCondition
	end if
end sub
%>

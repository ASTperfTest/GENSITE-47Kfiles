<%
sub xpCondition
	if request.form("htx_sTitle") <> "" then
		whereCondition = replace("ghtx.stitle LIKE N'%{0}%'", "{0}", request.form("htx_sTitle") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_iDept") <> "" then
		whereCondition = replace("ghtx.idept = N'{0}'", "{0}", request.form("htx_iDept") )
		fSql = fSql & " AND " & whereCondition
	end if
end sub
%>

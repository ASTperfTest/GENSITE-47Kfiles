<%
sub xpCondition
	if request.form("htx_ctUnitName") <> "" then
		whereCondition = replace("htx.ctUnitName LIKE N'%{0}%'", "{0}", request.form("htx_ctUnitName") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_CtUnitKind") <> "" then
		whereCondition = replace("htx.ctUnitKind LIKE N'%{0}%'", "{0}", request.form("htx_CtUnitKind") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_iBaseDSD") <> "" then
		whereCondition = replace("htx.ibaseDsd LIKE N'%{0}%'", "{0}", request.form("htx_iBaseDSD") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_inUse") <> "" then
		whereCondition = replace("htx.inUse LIKE N'%{0}%'", "{0}", request.form("htx_inUse") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_CheckYN") <> "" then
		whereCondition = replace("htx.checkYn LIKE N'%{0}%'", "{0}", request.form("htx_CheckYN") )
		fSql = fSql & " AND " & whereCondition
	end if
end sub
%>

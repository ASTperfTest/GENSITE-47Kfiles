<%
sub xpCondition
	if request.form("htx_sBaseTableName") <> "" then
		whereCondition = replace("htx.sBaseTableName LIKE N'%{0}%'", "{0}", request.form("htx_sBaseTableName") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_sBaseDSDName") <> "" then
		whereCondition = replace("htx.sBaseDSDName LIKE N'%{0}%'", "{0}", request.form("htx_sBaseDSDName") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_tDesc") <> "" then
		whereCondition = replace("htx.tDesc LIKE N'%{0}%'", "{0}", request.form("htx_tDesc") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_rdsdcat") <> "" then
		whereCondition = replace("htx.rdsdcat LIKE N'%{0}%'", "{0}", request.form("htx_rdsdcat") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_inUse") <> "" then
		whereCondition = replace("htx.inUse LIKE N'%{0}%'", "{0}", request.form("htx_inUse") )
		fSql = fSql & " AND " & whereCondition
	end if
end sub
%>

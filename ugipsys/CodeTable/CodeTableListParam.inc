<%
sub xpCondition
	if request.form("htx_codeID") <> "" then
		whereCondition = replace("htx.codeId LIKE N'%{0}%'", "{0}", request.form("htx_codeId") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_codeName") <> "" then
		whereCondition = replace("htx.codeName LIKE N'%{0}%'", "{0}", request.form("htx_codeName") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_codeTblName") <> "" then
		whereCondition = replace("htx.codeTblName LIKE N'%{0}%'", "{0}", request.form("htx_codeTblName") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_CodeType") <> "" then
		whereCondition = replace("htx.codeType LIKE N'%{0}%'", "{0}", request.form("htx_codeType") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_ShowOrNot") <> "" then
		whereCondition = replace("htx.showOrNot LIKE N'%{0}%'", "{0}", request.form("htx_showOrNot") )
		fSql = fSql & " AND " & whereCondition
	end if
end sub
%>

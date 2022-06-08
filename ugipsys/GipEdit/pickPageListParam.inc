<%
sub xpCondition
	if request.form("htx_iBaseDSD") <> "" then
		whereCondition = replace("htx.ibaseDsd = N'{0}'", "{0}", request.form("htx_iBaseDSD") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_CtRootID") <> "" then
		whereCondition = replace("n.ctRootId = N'{0}'", "{0}", request.form("htx_CtRootID") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_CatName") <> "" then
		whereCondition = replace("n.catName LIKE N'%{0}%'", "{0}", request.form("htx_CatName") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_CtUnitName") <> "" then
		whereCondition = replace("u.ctUnitName LIKE N'%{0}%'", "{0}", request.form("htx_CtUnitName") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_sTitle") <> "" then
		whereCondition = replace("htx.stitle LIKE N'%{0}%'", "{0}", request.form("htx_sTitle") )
		fSql = fSql & " AND " & whereCondition
	end if
end sub
	if request.form("htx_xKeyword") <> "" then
		whereCondition = replace("htx.xkeyword LIKE N'%{0}%'", "{0}", request.form("htx_xKeyword") )
		fSql = fSql & " AND " & whereCondition
	end if
%>

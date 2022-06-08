<%
sub xpCondition
	if request("htx_CtUnitID") <> "" then
		whereCondition = replace("htx.ictunit = {0}", "{0}", request("htx_CtUnitID") )
		fSql = fSql & " AND " & whereCondition
			SQL="Select ctUnitName from CtUnit where ctUnitId =" & request("htx_CtUnitID")
			SET RSS=conn.execute(SQL)
		xConStr = xConStr & "<tr><TD>主題單元：<td>" & RSS(0) & "</tr>"
	end if
	if request("htx_iBaseDSD") <> "" then
		whereCondition = replace("htx.ibaseDsd  = {0}", "{0}", request("htx_iBaseDSD") )
		fSql = fSql & " AND " & whereCondition
			SQL="Select sbaseDsdname from BaseDsd where ibaseDsd=" & request("htx_iBaseDSD")
			SET RSS=conn.execute(SQL)
		xConStr = xConStr & "<tr><TD>資料類別：<td>" & RSS(0) & "</tr>"
	end if
	if request("htx_IDateS") <> "" then
			rangeS = request("htx_IDateS")
			rangeE = request("htx_IDateE")
			if rangeE = "" then	rangeE=rangeS
			whereCondition = replace("htx.deditDate BETWEEN '{0}' and '{1}'", "{0}", rangeS)
			whereCondition = replace(whereCondition, "{1}", rangeE)
		fSql = fSql & " AND " & whereCondition
		xConStr = xConStr & "<tr><TD>日期範圍：<td>" & rangeS & "~" & rangeE & "</tr>"
	end if
	if request("htx_topCat") <> "" then
		whereCondition = replace("htx.topCat = '{0}'", "{0}", request("htx_topCat") )
		fSql = fSql & " AND " & whereCondition
			SQL="Select mvalue from CodeMain where codeMetaId='topDataCat' AND mcode=" _
				& pkStr(request("htx_topCat"),"")
			SET RSS=conn.execute(SQL)
		xConStr = xConStr & "<tr><TD>分網：<td>" & RSS(0) & "</tr>"
	end if
	if request("htx_idept") <> "" then
		whereCondition = replace("htx.idept = '{0}'", "{0}", request("htx_idept") )
		fSql = fSql & " AND " & whereCondition
			SQL="Select deptName from Dept WHERE deptId=" _
				& pkStr(request("htx_idept"),"")
			SET RSS=conn.execute(SQL)
		xConStr = xConStr & "<tr><TD>單位：<td>" & RSS(0) & "</tr>"
	end if
end sub
%>

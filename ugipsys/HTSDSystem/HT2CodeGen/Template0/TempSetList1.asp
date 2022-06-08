<% response.expires = 0 %>
<!--#Include file = "../inc/server_web.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
	if request.queryString("returnAP") <> "" then
		xstr = request.serverVariables("QUERY_STRING")
		xpos = instr(xstr,"&")
		if xpos>0 then		xqStr = mid(xstr,xpos+1)
		session("pickReturnAPURI")=request.queryString("returnAP")&".asp?"&xqStr
		pkCondition = ""
		setPKeyField = ""
		setPKeyValue = ""
		pkArray = split(xqStr,"&")
		if uBound(pkArray) > 0 then
		  for xi=0 to uBound(pkArray)
		  	xpstr = pkArray(xi)
			xpos = inStr(xpstr,"=")
			pkCondition = pkCondition & " AND " & left(xpstr,xpos-1) & "=" & pkStr(mid(xpstr,xpos+1),"")
			setPKeyField = setPKeyField & "," & left(xpstr,xpos-1)
			setPKeyValue = setPKeyValue & "," & pkStr(mid(xpstr,xpos+1),"")
		  next
		else
		  	xpstr = xqStr
			xpos = inStr(xpstr,"=")
			pkCondition = pkCondition & " AND " & left(xpstr,xpos-1) & "=" & pkStr(mid(xpstr,xpos+1),"")
			setPKeyField = setPKeyField & "," & left(xpstr,xpos-1)
			setPKeyValue = setPKeyValue & "," & pkStr(mid(xpstr,xpos+1),"")
		end if
		if len(pkCondition) > 0 then	pkCondition = mid(pkCondition,6)
		session("pkCondition") = pkCondition
		if len(setPKeyField) > 0 then	setPKeyField = mid(setPKeyField,2)
		session("setPKeyField") = setPKeyField
		if len(setPKeyValue) > 0 then	setPKeyValue = mid(setPKeyValue,2)
		session("setPKeyValue") = setPKeyValue
	end if

	if request.form("submitTask") = "½T¡@©w" then
		for each xx in request.form
'			response.write xx & "==>" & request(xx) & "<BR>"
		next
		pickArray=split(request.form("pickedItem"),", ")
		if uBound(pickArray) = 0 then	pickArray(0) = request.form("pickedItem")
		for xi=0 to uBound(pickArray)
'			response.write xi & "==>" & pickArray(xi) & "<BR>"

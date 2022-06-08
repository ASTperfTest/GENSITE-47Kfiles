<%
If session("ProgCode") <> HTProgCode Then
 session("ProgCode") = HTProgCode
 session("OrderList") = ""
 SetPageCon = session("PageCount")
ElseIf Request.Form("PageCount") <> "" Then
 SetPageCon = cint(Request.Form("PageCount"))
 session("PageCount") = cint(Request.Form("PageCount"))
ElseIf session("PageCount") <> "" Then
 SetPageCon = session("PageCount")
End If

If Request.Form("OrderList") <> "" Then
 session("OrderList") = Request.Form("OrderList")
End If

if request("page_no") = "" And request("NowPage") = "" then
  pno=0
elseIf request("NowPage") <> "" then
  pno=0
elseIf request("page_no") <> "" then
  pno=cint(request("page_no"))-1
end if

If request.form("SQL") <> "" Then
   session("querySQL") = request.form("SQL")
ElseIf request("page_no") = "" And Request.Form("PageCount") = "" And Request.Form("OrderList") = "" Then
   session("OrderList") = DefOrderList
   session("querySQL") = SetSQL
end if

 SQL = session("querySQL")

 If session("OrderList") <> "" Then SQL = SQL & " Order By "& session("OrderList")

'response.write SQL & "<hr>"
'response.end

	recno = 0
	OrderByCk = ""
  If SQL <> "" Then
	session("queryPage_no") = pno+1
	Set RS = Server.CreateObject("ADODB.RecordSet")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	RS.Open sql,Conn,1,1
Set RS = Conn.execute(sql)

'----------HyWeb GIP DB CONNECTION PATCH----------

	if not RS.eof then
		OrderByCk = "Y"
		recno = rs.recordcount
		pgno = int((recno-1)/SetPageCon)
		Pointer = pno*SetPageCon+1
		RS.absoluteposition=pointer
	End IF
   End If

  if pno * SetPageCon + SetPageCon > recno then
    tbi=recno mod SetPageCon
  else
    tbi=SetPageCon
  end if

 If OrderField <> "" And OrderName <> "" Then
	OrderField = Split(OrderField,",")
	OrderName = Split(OrderName,",")
	doSortDisplay = ""
		For XField = 0 to UBound(OrderField)
		  if xOrder = session("OrderList") then
		   doSortDisplay = doSortDisplay & "<option value="& OrderField(XField) &" selected>"& OrderName(XField) &"</option>" & vbcrlf
		  Else
		   doSortDisplay = doSortDisplay & "<option value="& OrderField(XField) &">"& OrderName(XField) &"</option>" & vbcrlf
		  End IF
		Next
 End IF
%>

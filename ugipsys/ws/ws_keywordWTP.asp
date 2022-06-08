<?xml version="1.0"  encoding="utf-8" ?>
<xKeywordList>
<%
FUNCTION pkStr (s, endchar)
  if s="" then
	pkStr = "null" & endchar
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	pkStr="'" & s & "'" & endchar
  end if
END FUNCTION
	xDate=cStr(date())
	xKeyword=pkStr(request.querystring("xKeyword"),"")
	Set conn=Server.CreateObject("ADODB.Connection")
       	Set Rs=Server.CreateObject("ADODB.Recordset")

'----------HyWeb GIP DB CONNECTION PATCH----------
'       	conn.Open session("ODBCDSN")
Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------

       	SQLCheck="Select iKeyword from CuDTKeywordWTP where iKeyword="+xKeyword
       	Set RSCheck=conn.execute(SQLCheck)
       	if not RSCheck.EOF then
       		SQLAct="Update CuDTKeywordWTP set iCount=iCount+1 where iKeyword="+xKeyword
       	else
       		SQLAct="Insert Into CuDTKeywordWTP values("+xKeyword+",1,null,'P','"+session("UserID")+"','"+xDate+"')"
       	end if
       	conn.execute(SQLAct)
	response.write "<xKeywordStr>Done</xKeywordStr>"       	
%>
</xKeywordList>

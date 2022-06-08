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
	xKeyword=pkStr(request.querystring("Corp"),"")
	Set conn=Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'       	conn.Open session("ODBCDSN")
Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------

    	    '----新增主表
    	    SQLIM="set nocount on;INSERT INTO CuDTGeneric(iBaseDSD,iCtUnit,fCTUPublic,iDept,sTitle,xPostDate) " & _
    		"Values(43,37,'N','0',"&xKeyword&",'"&date()&"'); select @@IDENTITY as NewID;"
    	    set RSx = conn.Execute(SQLIM)
    	    xNewIdentity = RSx(0)      
    	    '----新增演職人員表
    	    SQLIS="Insert Into CorpInformation(giCuItem) values("&xNewIdentity&")"
    	    conn.execute(SQLIS)	
	response.write "<xKeywordStr>Done</xKeywordStr>"       	
%>
</xKeywordList>

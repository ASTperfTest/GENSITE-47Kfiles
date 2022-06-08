<?xml version="1.0"  encoding="utf-8" ?>
<xKeywordList>
<%
function blen(xs)              '檢測中英文夾雜字串實際長度
  xl = len(xs)
  for p=1 to len(xs)
	if asc(mid(xs,p,1))<0 then xl = xl + 1
  next
  blen = xl
end function

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

Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'conn.Open session("ODBCDSN")
Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------


	xKeyword=pkStr(request.querystring("xKeyword"),"")
	xKeyword=mid(xKeyword,2)
	xKeyword=left(xKeyword,len(xKeyword)-1)
	xStr=""
	xReturnValue=""
	xKeywordArray=split(xkeyword,",")
	for i=0 to ubound(xKeywordArray)
		'----去除最後括號
		xPos=instrRev(xKeywordArray(i),"(")
		if xPos<>0 then
			xStr=Left(trim(xKeywordArray(i)),xPos-1)
		else
			xStr=trim(xKeywordArray(i))
		end if
		SQL="Select sTitle from ActorInformation AI Left Join CuDTGeneric CDT " & _
			" ON AI.giCuItem=CDT.iCuITem where sTitle=N'"&xStr&"'"
'			response.write "<SQL>"&SQL&"</SQL>"
		Set RSC=conn.execute(SQL)
		if RSC.eof then xReturnValue=xReturnValue+xStr+";"
	next
	if xReturnValue<>"" then xReturnValue=Left(xReturnValue,Len(xReturnValue)-1)
	response.write "<xKeywordStr>"+xReturnValue+"</xKeywordStr>"
%>
</xKeywordList>

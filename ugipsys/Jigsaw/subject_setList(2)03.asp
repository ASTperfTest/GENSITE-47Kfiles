<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" %>

<!--#include virtual = "/inc/server.inc" -->
<%

PageSize=request("PageSize")
iCUItem=request("iCUItem")
PageNumber=request("Page")
check=request("check")
uncheck=request("uncheck")
check1=session("jigcheck")
a=session("jigcheck1")
b=check
c=uncheck
checkarr = split(b, ";")
uncheckarr = split(c, ";")
'response.write request("gicuitem")

For i = 0 To UBound(checkarr)
   if (InStr(a,checkarr(i))>0) then
	
   else
	   add=checkarr(i)+";"
	   a=a+add
	end if
next

For i = 0 To UBound(uncheckarr)-1
    if (InStr(a,uncheckarr(i))>0) then
	   cut=uncheckarr(i)+";"
       a=Replace(a,cut,"")
    end if
next
session("jigcheck1")=a
response.write PageNumber 
'response.end
response.redirect "subject_setList(2).asp?iCUItem="&iCUItem&"&Page="&PageNumber&"&PageSize="&PageSize&"&gicuitem="&request("gicuitem")


%>
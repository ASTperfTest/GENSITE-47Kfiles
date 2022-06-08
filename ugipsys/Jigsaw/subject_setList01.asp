<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" %>

<!--#include virtual = "/inc/server.inc" -->
<%

CtRootId=request("CtRootId")
PageSize=request("PageSize")
iCUItem=request("iCUItem")
PageNumber=request("PageNumber")
check=request("check")
uncheck=request("uncheck")
check1=session("jigcheck")
a=session("jigcheck")
b=check
c=uncheck
checkarr = split(b, ";")
uncheckarr = split(c, ";")
response.write request("gicuitem")

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
session("jigcheck")=a
response.redirect "subject_setList.asp?iCUItem="&iCUItem&"&PageNumber="&PageNumber&"&PageSize="&PageSize&"&CtRootId="&CtRootId&"&gicuitem="&request("gicuitem")


%>
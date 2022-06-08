<%
Function CheckURL(URL)

URL=UCase(URL)

myFilterSymbols = "',"",<,>,`,!,$,^,*,\,(,),{,},[,],+,;,SCRIPT:"

myFilterSymbols=Split(myFilterSymbols,",")
errormes=now() & vbTab & "Error(1):REQUEST變數含特殊字元:"

For i = 0 To UBound(myFilterSymbols)
	

	
	if (InStr(URL,myFilterSymbols(i))>0) then
	  response.redirect "/"
	
	else
	   check=Server.URLEncode(myFilterSymbols(i))
			 if (InStr(URL,check)>0) then
				Log4U(errormes & myFilterSymbols(i) & vbTab & "URL:" & URL & vbTab & request.serverVariables("REMOTE_ADDR"))
		         response.redirect "/"
			 end if  
	end if
Next

End Function
%>

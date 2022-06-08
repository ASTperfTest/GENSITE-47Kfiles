<%
Function CheckURL(URL)

URL=UCase(URL)

myFilterSymbols = "',"",@,<,>,~,`,!,$,^,*,/,\,(,),{,},[,],-,+,;,SCRIPT:"

myFilterSymbols=Split(myFilterSymbols,",")


For i = 0 To UBound(myFilterSymbols)
	

	
	if (InStr(URL,myFilterSymbols(i))>0) then
	  response.redirect "/"
	
	else
	   check=hex(Asc(myFilterSymbols(i)))
	   check="%"&check
			 if (InStr(URL,check)>0) then
		         response.redirect "/"
			 end if  
	end if
Next

End Function
%>
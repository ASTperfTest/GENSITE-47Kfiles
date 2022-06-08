<%
'----940215新聞局電影網電影關聯_事業與演職人員關聯處理
function FilmRelated(xfunc,xTable,xType,xiCuItem,xFieldStr)
	Set Conn2 = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	Conn2.Open session("ODBCDSN")
'Set Conn2 = Server.CreateObject("HyWebDB3.dbExecute")
Conn2.ConnectionString = session("ODBCDSN")
Conn2.ConnectionTimeout=0
Conn2.CursorLocation = 3
Conn2.open
'----------HyWeb GIP DB CONNECTION PATCH----------

    	xxiCuItem=xiCuItem
    	if xTable="Corp" then 		'----Corp處理
	    if xfunc="edit" then 
	    	SQLD="delete from FilmCorpInfo where FilmNo="&xxiCuItem&" and CompanyType=N'"&xType&"'"  
	    	conn2.execute(SQLD)
	    end if
	    xKeywordArray=split(xFieldStr,",")
	    for i=0 to ubound(xKeywordArray)
	        xStr=trim(xKeywordArray(i))
		SQL="Select giCuItem from CorpInformation AI Left Join CuDTGeneric CDT " & _
			" ON AI.giCuItem=CDT.iCuITem where sTitle=N'"&xStr&"'"
		Set RSC=conn2.execute(SQL)
		if not RSC.eof then 
			SQLI=SQLI+"Insert Into FilmCorpInfo values("&xxiCuItem&","&RSC(0)&",'"&xType&"',null,null,null);"
		end if
	    next     
    	elseif xTable="Actor" then 	'----People處理
	    if xfunc="edit" then 
	    	SQLD="delete from FilmPeopleInfo where FilmNo="&xxiCuItem&" and RoleInfo=N'"&xType&"'"  
	    	conn2.execute(SQLD)
	    end if
	    xKeywordArray=split(xFieldStr,",")
	    for i=0 to ubound(xKeywordArray)
		'----取最後括號
		xPos=instrRev(xKeywordArray(i),"(")
		if xPos<>0 then
			xStr=Left(trim(xKeywordArray(i)),xPos-1)
			xStrPar="'"+mid(trim(xKeywordArray(i)),xPos)+"'"
		else
			xStr=trim(xKeywordArray(i))
			xStrPAr="null"
		end if
		SQL="Select giCuItem from ActorInformation AI Left Join CuDTGeneric CDT " & _
			" ON AI.giCuItem=CDT.iCuITem where sTitle=N'"&xStr&"'"
		Set RSC=conn2.execute(SQL)
		if not RSC.eof then 
			SQLI=SQLI+"Insert Into FilmPeopleInfo values("&xxiCuItem&","&RSC(0)&",'"&xType&"',"&xStrPAr&",null,null,null);"
		end if
	    next 
    	end if	
	if SQLI<>"" then conn2.execute(SQLI)  
'    	conn2.close
    	FilmRelated = ""
end function
%>


<%
haveBody = false
'小百科字辭及時解釋(取代字串)
function ReplaceAndFindKeyword(document)
    xs = document
	
	reviewmode = review
	'response.write xs
	xdd = ""
	if xs="" OR isNull(xs) then
  	ReplaceAndFindKeyword=""
  	exit function
    end if
	dim DomainUrl
	DomainUrl = "http://" & Request.ServerVariables("HTTP_HOST") &  Request.ServerVariables("URL")
	
	'replace 後台部分
	'判斷來源url
	if right(DomainUrl,10) = PhraseModUrl then
	
	fSql_new = "set nocount on "
    fSql_new = fSql_new & "declare  @ctRootId int set @ctRootId=" & request("mp")
    fSql_new = fSql_new & " DECLARE @titleList TABLE "
    fSql_new = fSql_new & "( "
    fSql_new = fSql_new & "titleId int "
    fSql_new = fSql_new & ",title nvarchar(200) "
    fSql_new = fSql_new & ",tType int default(0)"
    fSql_new = fSql_new & ",titletheme nvarchar(200) "
    fSql_new = fSql_new & ",rowId int "
    fSql_new = fSql_new & ") "
    fSql_new = fSql_new & "insert into @titleList (title) "
    fSql_new = fSql_new & "SELECT CuDTGeneric.sTitle FROM CuDTGeneric "
    fSql_new = fSql_new & "INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
    fSql_new = fSql_new & "WHERE (CuDTGeneric.iCTUnit = '2155') AND (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') "
    fSql_new = fSql_new & "union "
    fSql_new = fSql_new & "select  Phrase from SubjectPhrase where ctRootId = @ctRootId "
    fSql_new = fSql_new & "update rTitle "
    fSql_new = fSql_new & "set "
    fSql_new = fSql_new & "  rTitle.titleId = CuDTGeneric.iCUItem "
    fSql_new = fSql_new & " ,rTitle.tType = rTitle.tType | 1 "
    fSql_new = fSql_new & "FROM CuDTGeneric "
    fSql_new = fSql_new & "INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
    fSql_new = fSql_new & "inner join @titleList rTitle on rTitle.Title = CuDTGeneric.sTitle "
    fSql_new = fSql_new & "WHERE (CuDTGeneric.iCTUnit = '2155') AND (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') "
    fSql_new = fSql_new & "update rTitle "
    fSql_new = fSql_new & "set "
    fSql_new = fSql_new & "rTitle.titletheme= SubjectPhrase.Phrase "
    fSql_new = fSql_new & " ,rTitle.rowId = SubjectPhrase.rowId "
    fSql_new = fSql_new & " ,rTitle.tType = rTitle.tType | 2 "
    fSql_new = fSql_new & "FROM SubjectPhrase "
    fSql_new = fSql_new & "inner join @titleList rTitle on rTitle.Title = SubjectPhrase.Phrase "
    fSql_new = fSql_new & "where ctRootId =" & request("mp")
    fSql_new = fSql_new & " select * from @titleList "
	
	Set pedia = conn.execute(fSql_new)
 
	while not pedia.eof 
	    if trim(xs)=trim(pedia("title")) then   
			if IsNull(pedia("titleId"))  then
				xs = "<a id='ine_" & pedia("rowId") & "' class='jTip' name='/Phrase/ShowPhrase.asp?width=520&iCUItem=&themeId=" &  pedia("rowId") & "&mp=" & request("mp") & "&review=1'><span>view</span></a>"				
			else
				xs = "<a id='ine_" & pedia("titleId") & "' class='jTip' name='/Phrase/ShowPhrase.asp?width=520&iCUItem=" &  pedia("titleId") & "&themeId=" &  pedia("rowId") & "&mp=" & request("mp") & "&review=1'><span>view</span></a>"
			end if
		end if
		pedia.movenext	
	wend
	ReplaceAndFindKeyword = xs 
	
	
	'replace 前台部分
	else
	
	'session("PediaCtUnitId")
				fSql_new =  "set nocount on "
                fSql_new = fSql_new & "declare  @ctRootId int set @ctRootId =" & request("mp")
                fSql_new = fSql_new & " DECLARE @titleList TABLE "
                fSql_new = fSql_new & vbcrlf & "( "
                fSql_new = fSql_new & vbcrlf & "    titleId int "
                fSql_new = fSql_new & vbcrlf & "    ,title nvarchar(200) "
                fSql_new = fSql_new & vbcrlf & "    ,tType int default(0)"
                fSql_new = fSql_new & vbcrlf & "    ,titletheme nvarchar(200)"
                fSql_new = fSql_new & vbcrlf & "    ,rowId int"
				fSql_new = fSql_new & vbcrlf & ") "
                fSql_new = fSql_new & vbcrlf & "insert into @titleList (title) "
                fSql_new = fSql_new & vbcrlf & "SELECT CuDTGeneric.sTitle FROM CuDTGeneric " 
                fSql_new = fSql_new & vbcrlf & "INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
                fSql_new = fSql_new & vbcrlf & "WHERE (CuDTGeneric.iCTUnit = '2155') AND (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') "
                fSql_new = fSql_new & vbcrlf & "union "
                fSql_new = fSql_new & vbcrlf & "select  Phrase from SubjectPhrase where fCTUPublic = 'Y'  and ctRootId = @ctRootId "
                
                fSql_new = fSql_new & vbcrlf & "update rTitle "
                fSql_new = fSql_new & vbcrlf & "set "
                fSql_new = fSql_new & vbcrlf & "     rTitle.titleId = CuDTGeneric.iCUItem "
                fSql_new = fSql_new & vbcrlf & "    ,rTitle.tType = rTitle.tType | 1 "
                fSql_new = fSql_new & vbcrlf & "FROM CuDTGeneric "
                fSql_new = fSql_new & vbcrlf & "INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
                fSql_new = fSql_new & vbcrlf & "inner join @titleList rTitle on rTitle.Title = CuDTGeneric.sTitle "
                fSql_new = fSql_new & vbcrlf & "WHERE (CuDTGeneric.iCTUnit = '2155') AND (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') "
                
                fSql_new = fSql_new & vbcrlf & "update rTitle"
                fSql_new = fSql_new & vbcrlf & "set "
                fSql_new = fSql_new & vbcrlf & "     rTitle.titletheme= SubjectPhrase.Phrase"
                fSql_new = fSql_new & vbcrlf & "    ,rTitle.rowId = SubjectPhrase.rowId"
                fSql_new = fSql_new & vbcrlf & "    ,rTitle.tType = rTitle.tType | 2 "
                fSql_new = fSql_new & vbcrlf & "FROM SubjectPhrase "
                fSql_new = fSql_new & vbcrlf & "inner join @titleList rTitle on rTitle.Title = SubjectPhrase.Phrase "
                fSql_new = fSql_new & vbcrlf & "where SubjectPhrase.fCTUPublic = 'Y' and ctRootId = @ctRootId"
                fSql_new = fSql_new & vbcrlf & "select * from @titleList"
	
	
	set pedia = conn.execute(fSql_new)
 
	while not pedia.eof 
	    if (InStr(xs,trim(pedia("title"))) <> 0 ) then   
			if IsNull(pedia("titleId"))  then
				
			 xs = replace(xs,trim(pedia("title")),"<a id='ine_" & pedia("rowId")   & "' class='jTip jTip" & pedia("tType") & "' name='/subject/ShowPhrase.asp?width=520&iCUItem=&themeId=" &  pedia("rowId") & "&mp=" & request("mp") & "'><span>" & pedia("title")& "</span></a>",1,1) 
				
			 else

			 xs = replace(xs,trim(pedia("title")),"<a id='ine_" & pedia("titleId") & "' class='jTip jTip" & pedia("tType") & "' name='/subject/ShowPhrase.asp?width=520&iCUItem=" &  pedia("titleId") & "&themeId=" &  pedia("rowId") & "&mp=" & request("mp") & "'><span>" & pedia("title")& "</span></a>",1,1)
			 
			end if
			 
		end if
		pedia.movenext
		
	wend
	
	
	ReplaceAndFindKeyword = xs
	
	end if
	
end function


%>

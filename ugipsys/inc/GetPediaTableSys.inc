<%
function GetPediaTable(itemid,mpid,themeid,review)

'主題館跑以下資訊
if mpid <> 1 and themeid <> "" then

    '主題館　館主說明
	if review = 1 then
	Phrase="SELECT * FROM SubjectPhrase where rowId=" & themeid & " and CtRootID=" & mpid 
	else
	Phrase="SELECT * FROM SubjectPhrase where rowId=" & themeid & " and CtRootID=" & mpid & " and fCTUPublic='Y'"
	end if
	
	set themePhrase = conn.execute(Phrase)
	
	if not themePhrase.eof then
	    tableHtm = "<div class=""fade"" width=""480"">"
	    tableHtm = tableHtm & "<ul class=""tabs"">"
	    tableHtm = tableHtm & "<li><a>館主說明</a></li>"      
	    tableHtm = tableHtm & "</ul>"
	    tableHtm = tableHtm & "<div class=""items"">"
	    tableHtm = tableHtm & "<div id=""item2"">"
	    tableHtm = tableHtm & "<table width=""480"">"
		tableHtm = tableHtm & "<tr><td>" & Replace(Server.HTMLEncode(themePhrase("Content")),vbcrlf,"<br />") & "</td></tr>"
		tableHtm = tableHtm & "</table>"
		tableHtm = tableHtm & "</div>"
		tableHtm = tableHtm & "</div>"
		tableHtm = tableHtm & "</div>"

	end if
	
end if

if itemid <> "" then

	'百科詞彙
	sqlpediaup = " SELECT CuDTGeneric.sTitle, CuDTGeneric.xBody, CuDTGeneric.xKeyword, CuDTGeneric.vGroup, Pedia.engTitle, " _
                & " Pedia.formalName, Pedia.localName FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem " _
                & " WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') AND (CuDTGeneric.iCUItem = '"& itemid &"')"
  
    set pediatab = conn.execute(sqlpediaup)

	if not pediatab.eof then
	    tableHtm = tableHtm & "<div class=""fade"" width=""480"">"
	    tableHtm = tableHtm & "<ul class=""tabs"">"
	    tableHtm = tableHtm & "<li><a>農業小百科</a></li>"      
	    tableHtm = tableHtm & "</ul>"
	    tableHtm = tableHtm & "<div class=""items"">"
	    tableHtm = tableHtm & "<div id=""item2"">"
	    tableHtm = tableHtm & "<table width=""480"">"
		tableHtm = tableHtm & "<tr><th width=""80"" style=""text-align: right"">詞目</th><td>" & pediatab("sTitle") & "</td></tr>"
		if pediatab("engTitle") = Null or pediatab("engTitle") = ""  then
		else
		tableHtm = tableHtm & "<tr><th width=""80"" style=""text-align: right"">英文詞目</th><td>"
		    tableHtm = tableHtm & pediatab("engTitle") &"</td></tr>"
		end if
'		
'		tableHtm = tableHtm & "<tr><th width=""80"" style=""text-align: right"">學名(中/英)</th><td>"
'		if pediatab("formalName") = Null or  pediatab("formalName") = "" then
'		    tableHtm = tableHtm & "&nbsp;" &"</td></tr>"
'		else
'		    tableHtm = tableHtm & pediatab("formalName") &"</td></tr>"
'		end if
'		tableHtm = tableHtm & "<tr><th width=""80"" style=""text-align: right"">俗名(中/英)</th><td>"
'		if pediatab("localName") = Null or  pediatab("localName") = "" then
'		    tableHtm = tableHtm & "&nbsp;" &"</td></tr>"
'		else
'		    tableHtm = tableHtm & pediatab("localName") &"</td></tr>"
'		end if
if pediatab("xBody") = Null or pediatab("xBody") = "" then
		else
        tableHtm = tableHtm & "<tr><th width=""80"" style=""text-align: right"">名詞釋義</th><td>"  & Replace(pediatab("xBody"),vbCrlf, "<br />") & "</td></tr>"
		haveBody = true
		End If
		If pediatab("xKeyword") = Null Or pediatab("xKeyword") = "" Then
            'tableHtm = tableHtm & "<tr><th width=""80"" style=""text-align: right"">相關詞</th><td>&nbsp;</td></tr>"
        Else
		  Dim  items 
		    items = Split(pediatab("xKeyword"),";")
		    tableHtm = tableHtm & "<tr><th width=""80"" style=""text-align: right"">相關詞</th><td>"
         For Each item In items
            If item = Null or item = "" Then
			  else
			  tableHtm = tableHtm & item & ","
            End If
          Next
		  tableHtm = tableHtm & "</td></tr>"
		end if
		
		tableHtm = tableHtm & "</table>"
		tableHtm = tableHtm & "<div style=""text-align:right; padding-right:25px;""><a target='_blank' href='" & session("myWWWSiteURL") & "/Pedia/PediaContent.aspx?AId=" & itemid & "'>more...</div>"
	end if		
		tableHtm = tableHtm & "</div>"        
		tableHtm = tableHtm & "</div>"
end if

		tableHtm = tableHtm & "</div>"
		
		GetPediaTable = tableHtm
end function
%>
<%
function GetPediaTable(itemid,mpid,themeid)

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
		    supply = true
		else
		    supply = false
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
		
		if supply then'Grace 若沒有名詞釋義 就加入補充說明
		    tableHtm = tableHtm & "<br/><B>補充解釋：</B>"
		    tableHtm = tableHtm & "<table width=""480"">"
			tableHtm = tableHtm & "<tr><th></th><th style=""text-align: center"">補充說明</th><th width=""80"" style=""text-align: center"">發表者</th></tr>"
		    supplementSql = "SELECT Top 3 "_
		                  & "  ROW_NUMBER() OVER(ORDER BY Pedia.commendTime DESC) AS rowId "_
						  & ", CuDTGeneric.iCUItem "_
						  & ", CuDTGeneric.xBody "_
						  & ", Member.realname, Member.nickname "_
						  & ", Pedia.commendTime "_
						  & "FROM CuDTGeneric "_
						  & "INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "_
						  & "INNER JOIN Member ON Pedia.memberId = Member.account "_
						  & "WHERE (CuDTGeneric.fCTUPublic = 'Y') "_
						  & "AND (Pedia.xStatus = 'Y') "_
						  & "AND (CuDTGeneric.iCUItem in (select gicuitem from Pedia where parentIcuitem ='" & itemid & "' )) "_
						  & "ORDER BY Pedia.commendTime DESC "
						  
		    set supplement = conn.execute(supplementSql)
			while not supplement.eof
			    tableHtm = tableHtm & "<tr><td width='3%'>" & Replace(supplement("rowId"),vbCrlf, "<br />") & "</td>"
				tableHtm = tableHtm & "<td><a target='_blank' href='/Pedia/PediaExplainContent.aspx?AId=" & supplement("iCUItem") & "&PAId=" & itemid & "'>" & Replace(supplement("xBody"),vbCrlf, "<br />") & "</a></td>"
				tableHtm = tableHtm & "<td>" & Replace(supplement("nickname"),vbCrlf, "<br />") & "</td></tr>"
				supplement.movenext
			wend
			tableHtm = tableHtm & "</table>"
		end if
		
		tableHtm = tableHtm & "<div style=""text-align:right; padding-right:25px;""><a target='_blank' href='/Pedia/PediaContent.aspx?AId=" & itemid & "'>more...</div>"
	end if
		tableHtm = tableHtm & "</div>"        
		tableHtm = tableHtm & "</div>"
end if

		tableHtm = tableHtm & "</div>"
		
		GetPediaTable = tableHtm
end function
%>
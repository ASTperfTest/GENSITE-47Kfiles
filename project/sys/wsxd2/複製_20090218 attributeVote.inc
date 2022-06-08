<%
	'---態度投票---
	response.write "<queryItems>"
	for each item in request.querystring
		response.write "<" & item & ">" & request.querystring(item) & "</" & item & ">" 
	next
	response.write "</queryItems>"
	
	response.write "<attributeVote>"
		Dim voteIctUnit : voteIctUnit = 2180
		Dim totalVote : totalVote = 0
		Dim voteA : voteA = 0
		Dim voteB : voteB = 0
		Dim voteC : voteC = 0
		Dim voteD : voteD = 0
		Dim voteE : voteE = 0
		sql = "SELECT COUNT(*) AS voteCount, topCat FROM CuDTGeneric WHERE iCtUnit = " & voteIctUnit & " " & _
					"AND refID = " & xItem & " AND fCTUPublic = 'Y' GROUP BY topCat"		
					
		set rs = conn.execute(sql)
		while not rs.eof 
			select case rs("topCat")
				case "A"
					voteA = rs("voteCount")
				case "B"
					voteB = rs("voteCount")
				case "C"
					voteC = rs("voteCount")
				case "D"
					voteD = rs("voteCount")
				case "E"
					voteE = rs("voteCount")
			end select		
			rs.MoveNext
		wend
		rs.close
		set rs = nothing
		totalVote = voteA + voteB + voteC + voteD + voteE
		response.write "<totalVote>" & totalVote & "</totalVote>"
		response.write "<voteA>" & voteA & "</voteA>"
		response.write "<voteB>" & voteB & "</voteB>"
		response.write "<voteC>" & voteC & "</voteC>"
		response.write "<voteD>" & voteD & "</voteD>"
		response.write "<voteE>" & voteE & "</voteE>"
		if totalVote > 0 then		
			response.write "<voteAPercent>" & Round(voteA / totalVote * 100, 2) & "</voteAPercent>"
			response.write "<voteBPercent>" & Round(voteB / totalVote * 100, 2) & "</voteBPercent>"
			response.write "<voteCPercent>" & Round(voteC / totalVote * 100, 2) & "</voteCPercent>"
			response.write "<voteDPercent>" & Round(voteD / totalVote * 100, 2) & "</voteDPercent>"
			response.write "<voteEPercent>" & Round(voteE / totalVote * 100, 2) & "</voteEPercent>"					
			sql = "SELECT iCuItem, xBody FROM CuDTGeneric WHERE iCtUnit = " & voteIctUnit & " " & _
						"AND refID = " & xItem & " AND CAST(xBody AS nvarchar) <> '' AND fCTUPublic = 'Y' ORDER BY xPostDate DESC"			
			set rs = conn.execute(sql)
			while not rs.eof 
				response.write "<article>"
				response.write "<icuitem>" & rs("iCuItem") & "</icuitem>"
				response.write "<xBody><![CDATA[" & rs("xBody") & "]]></xBody>"
				response.write "</article>"
				rs.MoveNext
			wend
		end if
	response.write "</attributeVote>"
	'---end of 態度投票---
%>	
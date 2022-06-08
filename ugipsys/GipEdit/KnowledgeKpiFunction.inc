<%
	
	'---傳入問題的icuitem---
	sub DeleteQuestion( item )
		Dim disitems
		sql = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.iEditor, CONVERT(varchar, CuDTGeneric.xPostDate, 111) AS xPostDate "
		sql = sql & "FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
		sql = sql & "WHERE (KnowledgeForum.gicuitem = " & item & ") AND (KnowledgeForum.Status = 'N') "
		sql = sql & "AND (CuDTGeneric.iCtUnit = 932) "
		set ars = conn.execute(sql)		
		'response.write sql & "<hr />"
		while not ars.eof
			CheckShareDetail ars("iEditor"), "shareAsk", ars("xPostDate"), ars("iCUItem")			
			disitems = ars("iCUItem")
			ars.movenext
		wend 
		ars.close
		set ars = nothing
		UpdateStatus disitems
	end sub
	
	'---傳入討論的icuitem---
	sub DeleteDiscuss( item )
		Dim disitems
		sql = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.iEditor, CONVERT(varchar, CuDTGeneric.xPostDate, 111) AS xPostDate "
		sql = sql & "FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
		sql = sql & "WHERE (KnowledgeForum.gicuitem = " & item & ") AND (KnowledgeForum.Status = 'N') "
		sql = sql & "AND (CuDTGeneric.iCtUnit = 933) "
		set ars = conn.execute(sql)		
		'response.write sql & "<hr />"
		'response.end
		while not ars.eof
			CheckShareDetail ars("iEditor"), "shareDiscuss", ars("xPostDate"), ars("iCUItem")			
			disitems = ars("iCUItem")
			ars.movenext
		wend 
		ars.close
		set ars = nothing		
		UpdateStatus disitems
	end sub

	'---刪除評價---傳入討論的icuitem---
	sub DeleteCommend( item )
		Dim disitems
		sql = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.iEditor, CONVERT(varchar, CuDTGeneric.xPostDate, 111) AS xPostDate "
		sql = sql & "FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
		sql = sql & "WHERE (KnowledgeForum.ParentIcuitem = " & item & ") AND (KnowledgeForum.Status = 'N') "
		sql = sql & "AND (CuDTGeneric.iCtUnit = 938) "
		set ars = conn.execute(sql)		
		'response.write sql & "<hr />"
		'response.end
		while not ars.eof
			CheckShareDetail ars("iEditor"), "shareCommend", ars("xPostDate"), ars("iCUItem")			
			disitems = disitems & ars("iCUItem") & ","
			ars.movenext
		wend 
		ars.close
		set ars = nothing
		disitem = split(disitems, ",")
		for each i in disitem			
			if i <> "" then				
				UpdateStatus i
			end if
		next
	end sub
	
	'---刪除意見---傳入討論的icuitem---
	sub DeleteOpinion( item )
		Dim disitems
		sql = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.iEditor, CONVERT(varchar, CuDTGeneric.xPostDate, 111) AS xPostDate "
		sql = sql & "FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
		sql = sql & "WHERE (KnowledgeForum.ParentIcuitem = " & item & ") AND (KnowledgeForum.Status = 'N') "
		sql = sql & "AND (CuDTGeneric.iCtUnit = 935) "
		set ars = conn.execute(sql)		
		'response.write sql & "<hr />"
		while not ars.eof
			CheckShareDetail ars("iEditor"), "shareOpinion", ars("xPostDate"), ars("iCUItem")			
			disitems = disitems & ars("iCUItem") & ","
			ars.movenext
		wend 
		ars.close
		set ars = nothing
		disitem = split(disitems, ",")
		for each i in disitem
			if i <> "" then
				UpdateStatus i
			end if
		next
	end sub
	
	'---檢查shareDetail的內容---
	sub CheckShareDetail( memberid, sharetype, sharedate, articleid )
	
		Dim limitScore : limitScore = GetLimitScore()		
		Dim grade : grade = 0
		Dim oldGrade : oldGrade = 0
		Dim totalGrade : totalGrade = 0	'---當日最高得分---
		
		'---取出當日最後一筆---會得到最高分數---
		sql = "SELECT SUM(shareGrade) AS shareGrade FROM MemberGradeShareDetail "
		sql = sql & "WHERE memberId = '" & memberid & "' AND CONVERT(varchar, shareDate, 111) = '" & sharedate & "' "
		sql = sql & "AND shareSave <> 'N'"
		set srs = conn.execute(sql)		
		'response.write sql & "<hr />"
		if not srs.eof then
			totalGrade = srs("shareGrade")
		end if
		srs.close
		set srs = nothing 
		
		'---取出當筆的分數---
		sql = "SELECT memberId, shareDate, shareType, shareGrade, shareOldGrade, shareSave, articleId "
		sql = sql & "FROM MemberGradeShareDetail WHERE articleId = " & articleid
		set srs = conn.execute(sql)		
		'response.write sql & "<hr />"
		if not srs.eof then
			grade = srs("shareGrade")
			oldGrade = srs("shareOldGrade")
		end if
		srs.close
		set srs = nothing
		
		'---mark detail shareSave to 'N'
		MarkshareSave articleid
		
		'---減去分數---		
		if (totalGrade - grade) < limitScore then
			MinusGradeByTypeAndGrade memberid, sharedate, sharetype, grade		
		end if
				
	end sub
	
	sub MinusGradeByTypeAndGrade( memberid, sharedate, sharetype, grade )
		sql = "UPDATE MemberGradeShare SET " & sharetype & " = " & sharetype & " - " & grade & " "
		sql = sql & "WHERE memberId = '" & memberid & "' AND CONVERT(varchar, shareDate, 111) = '" & sharedate & "' "		
		conn.execute(sql)		
		'response.write sql & "<hr />"
	end sub
	
	sub MarkshareSave( articleid )
		sql = "UPDATE MemberGradeShareDetail SET shareSave = 'N' WHERE articleId = " & articleid		
		conn.execute(sql)		
		'response.write sql & "<hr />"
	end sub
	
	function GetLimitScore
		Dim limit
		sql = "SELECT Rank0_1 FROM kpi_set_score WHERE Rank0_2 = 'st_411'"
		set lrs = conn.execute(sql)
		if not lrs.eof then
			limit = lrs("Rank0_1")
		end if
		lrs.close
		set lrs = nothing
		GetLimitScore = limit
	end function
	
	sub UpdateStatus( articleid )
		sql = "UPDATE KnowledgeForum SET Status = 'D' WHERE gicuitem = " & articleid		
		conn.execute(sql)			
	end sub
	
%>

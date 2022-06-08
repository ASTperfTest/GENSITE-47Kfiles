<%
sub ActivityKnowledge(cuitem)
	'step 1:檢查是否為活動期間
	'step 2:判斷得分種類 1:一般發問  2分 2:主題館發問 3分 4:已存在討論 1分
	'step 3:新增發問得分至活動資料表
	'step 3-1: 新增前先判斷是否已存在資料表中,避免重複作業
	'step 1------------------------------------------------------------
	
	if ActivityEndFlag() = "true" then
		'判斷是否已存在活動資料表中
		sql = "SELECT CUItemId FROM dbo.KnowledgeActivity WHERE CUItemId = '" & cuitem &"'"
		set rs = conn.execute(sql)
		if not rs.eof then
			sql2 = "UPDATE KnowledgeActivity SET State=1,ModifyTime=GETDATE()  WHERE CUItemId = " & cuitem
			conn.execute(sql2)
			exit sub
		end if
		rs.close
		set rs = nothing
		
		'step2--------------------------------------------------------
		Dim score,strType,editor
		sql = "SELECT iCUItem FROM dbo.ActivitySubjectKnowledge WHERE iCUItem = '" & cuitem &"'"
		set rs = conn.execute(sql)
		if not rs.eof then
			score = "3"
			strType = "2"
		else
			score = "2"
			strType = "1"
		end if
		rs.close
		set rs = nothing
		'新增發問得分
		'取得題目資訊
		sql = "SELECT iEditor FROM dbo.CuDTGeneric WHERE iCUItem = '" & cuitem &"'"
		set rs = conn.execute(sql)
		if not rs.eof then
			editor = rs("iEditor")
		else
			editor = ""
		end if
		rs.close
		set rs = nothing
		
		sql2 = "INSERT INTO KnowledgeActivity (CUItemId ,Type ,Grade ,MemberId ,State) VALUES(" & cuitem &"," & strType &", " & score &", '" & editor &"', 1) "
		conn.execute(sql2)

		'------檢查是否有討論------
		sql = "SELECT CuDTGeneric.iCUItem,CONVERT(VARCHAR(20),xPostDate,120) AS xPostDate ,iEditor FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
		sql = sql & "WHERE (KnowledgeForum.ParentIcuitem = " & cuitem & ") AND (KnowledgeForum.Status = 'N') AND (CuDTGeneric.iCtUnit = 933) AND ( CuDTGeneric.fCTUPublic = 'Y' ) AND ( CuDTGeneric.siteId = 3 )"

		set rs = conn.execute(sql)
		dim discussScore ,checkTime
		Dim activityId 
		activityId = GetActivityId()
		while not rs.eof 
			''檢查討論發表時間:活動前得1分,活動中得2分
			checkTime = rs("xPostDate")
			sql5 = "SELECT ActivityId FROM Activity WHERE ( CONVERT(DATETIME,'" & checkTime & "') BETWEEN ActivityStartTime AND ActivityEndTime) AND ActivityId = '" & activityId & "'"
	
			set rs5 = conn.execute(sql5)
			if not rs5.eof then
				
				discussScore = "2"
			else
				discussScore = "1"
			
			end if
			rs5.close
			set rs5 = nothing

			sql2 = "INSERT INTO KnowledgeActivity (CUItemId ,Type ,Grade ,MemberId ,State) VALUES(" & rs("iCUItem") &",4, " & discussScore & ", '" & rs("iEditor") &"', 1) "
			conn.execute(sql2)		
			rs.movenext
		wend
		rs.close
		set rs = nothing
		
	else
		exit sub
	end if
	

end sub

sub PublicTempDiscuss(cuitem)

'活動結束就不再做任何處理
	if ActivityEndFlag() = "false" then
		exit sub
	else
		sql = "SELECT CUItemId FROM dbo.KnowledgeActivity WHERE CUItemId = '" & cuitem &"'"
		set rs = conn.execute(sql)
		if rs.eof then
			exit sub
		end if
		rs.close
		set rs = nothing
		
		'找出暫存討論
		sql = "SELECT CuDTGeneric.iCUItem,CONVERT(VARCHAR(20),xPostDate,120) AS xPostDate ,iEditor FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
		sql = sql & "WHERE (KnowledgeForum.ParentIcuitem = " & cuitem & ") AND (KnowledgeForum.Status = 'N') AND (CuDTGeneric.iCtUnit = 933) AND ( CuDTGeneric.fCTUPublic = 'N' ) AND ( CuDTGeneric.siteId = 3 ) AND ( ISNULL(CuDTGeneric.Abstract,'') <> 'Y')"

		set rs = conn.execute(sql)
		
		dim discussScore ,checkTime,activityId
		activityId = GetActivityId()
		
		while not rs.eof 
			'公開所有暫存討論,並算活動積分
			sqlUpd = "UPDATE CuDTGeneric SET fCTUPublic='Y' WHERE iCUItem = " & rs("iCUItem")
			conn.execute(sqlUpd)
			
			'更新討論數
			sqlUdpDis = "UPDATE KnowledgeForum SET DiscussCount = DiscussCount +1 WHERE gicuitem = " & cuitem
			conn.execute(sqlUdpDis)
			
			'紀錄公開的討論
			sqlIns = "INSERT INTO ActivityTempDiscuss ( iCUItem, MemberId, PublishTime ) VALUES( " & rs("iCUItem") & ", '" & rs("iEditor") & "', GETDATE())"
			conn.execute(sqlIns)
			
			''檢查討論發表時間:活動前得1分,活動中得2分
			checkTime = rs("xPostDate")
			sql5 = "SELECT ActivityId FROM Activity WHERE ( CONVERT(DATETIME,'" & checkTime & "') BETWEEN ActivityStartTime AND ActivityEndTime) AND ActivityId = '" & activityId & "'"
	
			set rs5 = conn.execute(sql5)
			if not rs5.eof then
				discussScore = "2"
			else
				discussScore = "1"
			
			end if
			rs5.close
			set rs5 = nothing
			'計算活動得分
			sql2 = "INSERT INTO KnowledgeActivity (CUItemId ,Type ,Grade ,MemberId ,State) VALUES(" & rs("iCUItem") &",4, " & discussScore & ", '" & rs("iEditor") &"', 1) "
			conn.execute(sql2)		
			rs.movenext
		wend
		rs.close
		set rs = nothing

	end if

end sub


sub UnusableActivityKnowledge(cuitem)
	
	'活動結束就不再做任何處理
	if ActivityEndFlag() = "false" then
		exit sub
	else
		'判斷是否為活動題目
		sql = "SELECT CUItemId FROM dbo.KnowledgeActivity WHERE CUItemId = '" & cuitem &"'"
		set rs = conn.execute(sql)
		if rs.eof then
			exit sub
		end if
		rs.close
		set rs = nothing
		
		'State=0 表示此活動下架、不公開 or 刪除
		sql2 = "UPDATE KnowledgeActivity SET State=0,ModifyTime=GETDATE()  WHERE CUItemId = " & cuitem
		conn.execute(sql2)
	end if

end sub

sub ActivityDiscuss(cuitem,add)

	'判斷是否為活論討論
	sql = "SELECT CUItemId FROM dbo.KnowledgeActivity WHERE CUItemId = '" & cuitem &"'"
	set rs = conn.execute(sql)
	if rs.eof then
		exit sub
	end if
	rs.close
	set rs = nothing
	
	'State=0 表示此討論不納入活動討論中,反之則納入
	if add = "false" then
		sql2 = "UPDATE KnowledgeActivity SET State=0,ModifyTime=GETDATE()  WHERE CUItemId = " & cuitem
	else
		sql2 = "UPDATE KnowledgeActivity SET State=1,ModifyTime=GETDATE()  WHERE CUItemId = " & cuitem
	end if
	conn.execute(sql2)
	
end sub



function ActivityFlag

	ActivityFlag= "false"
	Dim activityId 
	activityId = GetActivityId()	

	sql = "SELECT ActivityId FROM Activity WHERE (GETDATE() BETWEEN ActivityStartTime AND ActivityEndTime) AND ActivityId = '" & activityId & "'"
	set rs = conn.execute(sql)
	if not rs.eof then
		ActivityFlag = "true"
	end if
	rs.close
	set rs = nothing

end function




function ActivityEndFlag

	ActivityEndFlag= "false"
	Dim activityId 
	activityId = GetActivityId()

	sql = "SELECT ActivityId FROM Activity WHERE (GETDATE() < ActivityEndTime) AND ActivityId = '" & activityId & "'"
	set rs = conn.execute(sql)
	if not rs.eof then
		ActivityEndFlag = "true"
	end if
	rs.close
	set rs = nothing

end function

Function GetActivityId

	GetActivityId= "knowledge2010" 
	
End Function



'管理者手動將問題反餽回主題館的額外積分(得1分)
sub PublishToSubject(cuitem)

	sql = "SELECT CUItemId FROM dbo.KnowledgeActivity WHERE CUItemId = '" & cuitem &"' AND Grade < 3" 
		
	set rs = conn.execute(sql)
	if rs.eof then
		exit sub
	end if
	rs.close
	set rs = nothing
			
	sql2 = "UPDATE KnowledgeActivity SET Grade=Grade +1 ,ModifyTime=GETDATE()  WHERE CUItemId = " & cuitem
	conn.execute(sql2)
	
end sub

%>
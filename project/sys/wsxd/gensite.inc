<%

	'---主題館加上會員---20080916---vincent---
	
	response.write "<login>"
		Dim memName, memnickname, memlogincount, memAccount
		if request("memID") <> "" Then		
			sql = "SELECT realname, ISNULL(nickname, '') AS nickname, ISNULL(logincount, 0) AS logincount,Account " & _
						"FROM Member where account = '" & request("memID") & "'"
			Set memrs = conn.Execute(sql)
			If Not memrs.EOF Then
				response.write "<status>true</status>"	
				memName = trim(memrs("realname"))
				memnickname = trim(memrs("nickname"))								
				memlogincount = memrs("logincount")
				if memnickname = "" then
					memAccount =memrs("Account")
				else
					memAccount =memrs("nickname")
				end if
			else 
					response.write "<status>false</status>"	
			End If
		else
			response.write "<status>false</status>"
		End If
		response.write "<memID>" & request("memID") & "</memID>"
		response.write "<memName><![CDATA[" & memName & "]]></memName>"
		response.write "<memNickName><![CDATA[" & memnickname & "]]></memNickName>"
		response.write "<memLoginCount>" & memlogincount & "</memLoginCount>"		
		response.write "<memAccount>" & memAccount & "</memAccount>"	
		response.write "<gstyle>" & request("gstyle") & "</gstyle>"				
	response.write "</login>"

	'---鎖右鍵-20080916---vincent---
	sql = "SELECT * FROM NodeInfo WHERE CtrootID = " & xRootID
	set nrs = conn.execute(sql)
	response.write "<lockrightbtn>"
	if not nrs.eof then
		if nrs("lockRightBtn") = "Y" then 
			response.write "Y"
		else
			response.write "N"
		end if
	end if
	response.write "</lockrightbtn>"
	nrs.close
	set nrs = nothing

	sqlb = "SELECT ctRootName FROM CatTreeRoot Where ctRootId=" & xRootID
	Set rsb = conn.execute(sqlb)
	response.write "<rootname>" & rsb(0)& "</rootname>"
	
	
	sql = "SELECT isnull(pic_title,'') as pic_title, isnull(pic_banner,'') as pic_banner, list, data FROM nodeinfo where ctrootid ='" & xRootID & "'"
	Set rss = conn.Execute(sql)
	If Not rss.EOF Then
		layoutStyle99 = "mp" & right(trim(rss("list")),1)
		cssStyle99 = trim(rss("data"))
		pic_banner = trim(rss("pic_banner"))
		pic_title = trim(rss("pic_title"))
		'橫幅圖片
		if pic_banner <> "" then
			if instr(lcase(pic_banner),"customerbanner")>0 then
				response.Write "<BannerPicFilePath>public/"& pic_banner &"</BannerPicFilePath>"
				response.Write "<BannerPicFileName>主題館橫幅</BannerPicFileName>"
			elseif instr(lcase(pic_banner),"banner")>0 then				
				pic_banner_ary = split(pic_banner,"-")
				response.Write "<BannerPicFilePath>xslgip/" & cssStyle99 &"/images/"& pic_banner & "</BannerPicFilePath>"
				response.Write "<BannerPicFileName>" & pic_banner_ary(1) & "</BannerPicFileName>"
			end if
		end if
		'標題圖片
		if pic_title <>"" then
			response.Write "<TitlePicFilePath>public/" & pic_title & "</TitlePicFilePath>"
			response.Write "<TitlePicFileName>主題館logo</TitlePicFileName>"
		else
			response.Write "<TitlePicFilePath>xslgip/" & cssStyle99 & "/images/logo.gif</TitlePicFilePath>"
			response.Write "<TitlePicFileName>農委會茄子館</TitlePicFileName>"
		end if
			
	Else
		layoutStyle99 = "mp"
		cssStyle99 = ""
		banner_path = "banner-農產橫幅-57635.jpg"
		banner_name = "農產橫幅"
		picTitle_path = "logo.gif"
		picTitle_name = "農委會logo"
		response.Write "<BannerPicFilePath>" & banner_path & "</BannerPicFilePath>"
		response.Write "<BannerPicFileName>" & banner_name & "</BannerPicFileName>"
		response.Write "<TitlePicFilePath>" & picTitle_path & "</TitlePicFilePath>"
		response.Write "<TitlePicFileName>" & picTitle_name & "</TitlePicFileName>"
	End If
	
	response.Write "<layoutStyle>" & layoutStyle99 & "</layoutStyle>"
	if cssStyle99<>"" then
		response.Write "<MpStyle>" & cssStyle99 & "</MpStyle>"
	else
		response.Write "<MpStyle>" & nullText(refModel.selectSingleNode("MpStyle")) & "</MpStyle>"
	end if
	'---主題館關鍵字-2008/12/27---vincent--- 先看有沒有設定沒有在抓系統計算的
	response.write "<BlockKW>"
	response.write "<Caption>主題館關鍵字</Caption>"
	sql = "SELECT Keywords FROM NodeInfo WHERE CtrootID = " & xRootID
	set keywordrs = conn.execute(sql)
	if not keywordrs.eof then
	    if keywordrs("Keywords") <> "" or keywordrs("Keywords") <> NULL then
	        Dim kyArray
		    Dim kylength
		    kyArray = Split(keywordrs("Keywords"), ",")
		    kylength = ubound(kyArray)
		    for i = 0 to kylength
		        response.write "<Article><Caption><![CDATA[" & kyArray(i) & "]]></Caption></Article>"
		    next
			else
			keywordrs.close
	        set keywordrs = nothing
	        sql = "SELECT TOP 10 KEYWORD FROM REPORT_KEYWORD_FREQUENCY " & _
				"WHERE MP = " & request("mp") & " GROUP BY KEYWORD ORDER BY SUM(FREQUENCY) DESC "
	        set keywordrs = KMConn.execute(sql)
	        while not keywordrs.eof
		        response.write "<Article><Caption><![CDATA[" & keywordrs("KEYWORD") & "]]></Caption></Article>"
		        keywordrs.movenext
	        wend
		end if
	else
	    keywordrs.close
	    set keywordrs = nothing
	    sql = "SELECT TOP 10 KEYWORD FROM REPORT_KEYWORD_FREQUENCY " & _
				"WHERE MP = " & request("mp") & " GROUP BY KEYWORD ORDER BY SUM(FREQUENCY) DESC "
	    set keywordrs = KMConn.execute(sql)
	    while not keywordrs.eof
		    response.write "<Article><Caption><![CDATA[" & keywordrs("KEYWORD") & "]]></Caption></Article>"
		    keywordrs.movenext
	    wend
	end if
	keywordrs.close
	set keywordrs = nothing
	response.write "</BlockKW>"
%>
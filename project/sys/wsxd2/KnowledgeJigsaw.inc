<%
	'---知識拚圖---
	Dim jigsawictunit : jigsawictunit = 2199
	Dim jigsawFlag : jigsawFlag = false
	'---檢查傳入的xitem是否為知識拚圖的ictunit---
	jigsawsql = "SELECT * FROM CuDTGeneric WHERE iCUItem = " & xItem & " AND iCTUnit = " & jigsawictunit & " AND fCTUPublic = 'Y' "
	set jigsawrs = conn.execute(jigsawsql)
	if not jigsawrs.eof then
		jigsawFlag = true
	end if
	jigsawrs.close
	set jigsawrs = nothing
	
	if jigsawFlag then
		response.write "<jigsaw>"
		
		'---專區---
		jigsawsql = "SELECT * FROM CuDTGeneric WHERE icuitem = " & xItem							
		set jigsawrs = conn.execute(jigsawsql)
		if not jigsawrs.eof then
			response.write "<sTitle><![CDATA[" & jigsawrs("sTitle") & "]]></sTitle>"					
		end if
		jigsawrs.close
		set jigsawrs = nothing
		
		'---最新議題---
		GetLatestIssue xItem
		
		'---取出議題排序---
		Dim unit, subject, tank, home
		jigsawsql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.icuitem = KnowledgeJigsaw.gicuitem " _
							& "WHERE KnowledgeJigsaw.parenticuitem = " & xItem & " AND CuDTGeneric.topCat = 'C'"
		set jigsawrs = conn.execute(jigsawsql)
		if not jigsawrs.eof then
			unit = jigsawrs("orderSiteUnit")
			subject = jigsawrs("orderSubject")
			tank = jigsawrs("orderKnowledgeTank")
			home = jigsawrs("orderKnowledgeHome")
		end if
		jigsawrs.close
		set jigsawrs = nothing
		
		'---產生議題xml by order---
		GenArticleXml xItem, unit, subject, tank, home
		
		'---相關資源推薦---
		GetRefLinkXml xItem
		
		'---added by Joey 討論區---
		GetDiscussionXml xItem
		
		
		response.write "</jigsaw>"
	end if
	
	
	'added by Joey --議題討論區
	sub GetDiscussionXml(Id)
		response.write "<discussion>"
		jigid = 0
		flag = true
		jigsawsql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.icuitem = KnowledgeJigsaw.gicuitem " _
							& "WHERE KnowledgeJigsaw.parentIcuitem = " & id & " AND CuDTGeneric.topCat = 'F'"
		set jigsawrs = conn.execute(jigsawsql)
		
		if not jigsawrs.eof then 
			jigid = jigsawrs("icuitem")
			response.write "<isopen>" & jigsawrs("fCTUPublic") & "</isopen>"
		end if
		
		
		jigsawrs.close
		set jigsawrs = nothing
		'---列出文章---
		jigsql = "SELECT CuDTGeneric.xBody,CuDTGeneric.iEditor, KnowledgeJigsaw.path, convert(varchar, CuDTGeneric.xpostdate, 111) as xpostdate FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.iCUItem = KnowledgeJigsaw.gicuitem " & _
					   "WHERE (KnowledgeJigsaw.parentIcuitem = " & jigid & ") AND (KnowledgeJigsaw.Status = 'Y') " & _
						 "ORDER BY KnowledgeJigsaw.orderArticle DESC,xpostdate ASC"
		set jigrs = conn.execute(jigsql)		
		while not jigrs.eof 
			response.write "<article>"
			response.write "<xBody><![CDATA[" & jigrs("xBody") & "]]></xBody>"
			response.write "<iEditor><![CDATA[" & jigrs("iEditor") & "]]></iEditor>"
			response.write "<xpostdate>" & jigrs("xpostdate") & "</xpostdate>"			
			response.write "</article>"
			jigrs.movenext
		wend
		response.write "</discussion>"
		'by Joey 
		'response.write "<sqljoey1>" & jigsawsql & "</sqljoey1>"
		'response.write "<sqljoey2>" & jigsql & "</sqljoey2>"
	end sub	
	
	'---最新議題---	
	sub GetLatestIssue( id )
		response.write "<latestissue>"
		jigid = 0
		flag = true
		jigsawsql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.icuitem = KnowledgeJigsaw.gicuitem " _
							& "WHERE KnowledgeJigsaw.parentIcuitem = " & id & " AND CuDTGeneric.topCat = 'A'"
		set jigsawrs = conn.execute(jigsawsql)
		
		
		if not jigsawrs.eof then 
			jigid = jigsawrs("icuitem")
			response.write "<isopen>" & jigsawrs("fCTUPublic") & "</isopen>"
		end if
		jigsawrs.close
		set jigsawrs = nothing
		'---列出文章---
		jigsql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.iCUItem = KnowledgeJigsaw.gicuitem " & _
					   "WHERE (KnowledgeJigsaw.parentIcuitem = " & jigid & ") AND (KnowledgeJigsaw.Status = 'Y') " & _
						 "ORDER BY KnowledgeJigsaw.orderArticle DESC"
		set jigrs = conn.execute(jigsql)		
		while not jigrs.eof 
			response.write "<article>"
			if flag then
				GetLatestFirstArticle jigrs("ArticleId"), jigrs("CtRootId"), jigrs("path")
				flag = false
			else
				response.write "<stitle><![CDATA[" & jigrs("sTitle") & "]]></stitle>"
				response.write "<top>N</top>"
				response.write "<content></content>"
				response.write "<path><![CDATA[" & jigrs("path") & "]]></path>"
				response.write "<img><![CDATA[" & jigrs("xImgFile") & "]]></img>"
			end if
			jigrs.movenext
			response.write "</article>"
		wend
		response.write "</latestissue>"
		
		response.write "<sqlJoey1>" & jigsawsql & "</sqlJoey1>"
		response.write "<sqlJoey2>" & jigsql & "</sqlJoey2>"
	end sub
	
	sub GetLatestFirstArticle( aid, rid, path )		
		if rid = "4" then
			kmsql = "SELECT SUBJECT, DESCRIPTION FROM REPORT WHERE REPORT_ID = '" & aid & "'"
			set kmrs = KMConn.execute(kmsql)
			if not kmrs.eof then
				response.write "<stitle><![CDATA[" & kmrs("SUBJECT") & "]]></stitle>"
				response.write "<top>Y</top>"
				if len(kmrs("DESCRIPTION")) > 170 then
					response.write "<content><![CDATA[" & mid(kmrs("DESCRIPTION"), 1, 170) & "...]]></content>"
				else
					response.write "<content><![CDATA[" & kmrs("DESCRIPTION") & "...]]></content>"
				end if				
				response.write "<path><![CDATA[" & path & "]]></path>"
				response.write "<img></img>"
			else
				response.write "<stitle></stitle>"
				response.write "<content></content>"
				response.write "<path></path>"
				response.write "<img></img>"
			end if
			kmrs.close
			set kmrs = nothing
		else
			sql = "SELECT sTitle, CAST(xBody AS nvarchar(500)) AS xBody, xImgFile FROM CuDTGeneric WHERE icuitem = " & aid
			set newrs = conn.execute(sql)
			
			'response.write "<joey>" & sql & "</joey>"
			
			if not newrs.eof then
				response.write "<stitle><![CDATA[" & newrs("sTitle") & "]]></stitle>"
				response.write "<top>Y</top>"
				if len(newrs("xBody")) > 170 then
					response.write "<content><![CDATA[" & mid(newrs("xBody"), 1, 170) & "...]]></content>"
				else
					response.write "<content><![CDATA[" & newrs("xBody") & "...]]></content>"
				end if				
				response.write "<path><![CDATA[" & path & "]]></path>"
				response.write "<img><![CDATA[" & newrs("xImgFile") & "]]></img>"
			else
				response.write "<stitle></stitle>"
				response.write "<content></content>"
				response.write "<path></path>"
				response.write "<img></img>"
			end if
			newrs.close
			set newrs = nothing
		end if
	end sub
	
	'---產生 議題關聯知識文章 xml by order---	
	sub GenArticleXml( id, unit, subject, tank, home )
		response.write "<issuearticle>"
		response.write "<order>"
			response.write "<unit>" & unit & "</unit>"
			response.write "<subject>" & subject & "</subject>"
			response.write "<tank>" & tank & "</tank>"
			response.write "<home>" & home & "</home>"
		response.write "</order>"
		jigid = 0
		jigsawsql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.icuitem = KnowledgeJigsaw.gicuitem " _
							& "WHERE KnowledgeJigsaw.parentIcuitem = " & id & " AND CuDTGeneric.topCat = 'B'"
		set jigsawrs = conn.execute(jigsawsql)
		if not jigsawrs.eof then 
			jigid = jigsawrs("icuitem")
			response.write "<isopen>" & jigsawrs("fCTUPublic") & "</isopen>"
		end if
		jigsawrs.close
		set jigsawrs = nothing		
		
		GetArticleXml "unit", jigid, 1
		GetArticleXml "subject", jigid, 2
		GetArticleXml "home", jigid, 3
		GetArticleXml "tank", jigid, 4
		
		'//議題關聯影音
		jigid = 0
		jigsawsql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.icuitem = KnowledgeJigsaw.gicuitem " _
							& "WHERE KnowledgeJigsaw.parentIcuitem = " & id & " AND CuDTGeneric.topCat = 'D'"
		set jigsawrs = conn.execute(jigsawsql)
		if not jigsawrs.eof then 
			jigid = jigsawrs("icuitem")
			response.write "<isopen>" & jigsawrs("fCTUPublic") & "</isopen>"
		end if
		jigsawrs.close
		set jigsawrs = nothing
		GetArticleXml "media", jigid, 1
	
		response.write "</issuearticle>"
	end sub
	
	'相關資源推薦
	sub GetRefLinkXml(Id)
		response.write "<reflink>"
		jigid = 0
		flag = true
		jigsawsql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.icuitem = KnowledgeJigsaw.gicuitem " _
							& "WHERE KnowledgeJigsaw.parentIcuitem = " & id & " AND CuDTGeneric.topCat = 'E'"
		set jigsawrs = conn.execute(jigsawsql)
		
		if not jigsawrs.eof then 
			jigid = jigsawrs("icuitem")
			response.write "<isopen>" & jigsawrs("fCTUPublic") & "</isopen>"
		end if
		
		
		jigsawrs.close
		set jigsawrs = nothing
		'---列出文章---
		jigsql = "SELECT CuDTGeneric.stitle, KnowledgeJigsaw.path, convert(varchar, CuDTGeneric.xpostdate, 111) as xpostdate FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.iCUItem = KnowledgeJigsaw.gicuitem " & _
					   "WHERE (KnowledgeJigsaw.parentIcuitem = " & jigid & ") AND (KnowledgeJigsaw.Status = 'Y') " & _
						 "ORDER BY KnowledgeJigsaw.orderArticle DESC"
		set jigrs = conn.execute(jigsql)		
		while not jigrs.eof 
			response.write "<article>"
			response.write "<title><![CDATA[" & jigrs("stitle") & "]]></title>"
			response.write "<url><![CDATA[" & jigrs("path") & "]]></url>"
			response.write "<xpostdate>" & jigrs("xpostdate") & "</xpostdate>"			
			response.write "</article>"
			jigrs.movenext
		wend
		response.write "</reflink>"
		'by Joey 
		'response.write "<sqljoey1>" & jigsawsql & "</sqljoey1>"
		'response.write "<sqljoey2>" & jigsql & "</sqljoey2>"
	end sub	
	
	
	sub GetArticleXml( atype, jigid, rid )
			
		response.write "<" & atype & ">"
		if atype = "media" then
			Dim count2 : count2 = 0
			for ii =1 to 4
				'---列出文章---
				jigsql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.iCUItem = KnowledgeJigsaw.gicuitem " & _
								 "WHERE (KnowledgeJigsaw.parentIcuitem = " & jigid & ") AND (KnowledgeJigsaw.Status = 'Y') AND (KnowledgeJigsaw.CtRootId = " & ii & ") " & _
								 "ORDER BY KnowledgeJigsaw.orderArticle DESC"
				set jigrs = conn.execute(jigsql)
				'response.write "<sql>" & jigsql & "</sql>"
				while not jigrs.eof 
					response.write "<article>"				
					response.write "<stitle>" & jigrs("stitle") & "</stitle>"				
					response.write "<path><![CDATA[" & jigrs("path") & "]]></path>"
					GetIssueArticlePostDate jigrs("ArticleId"), jigrs("CtRootId")
					jigrs.movenext
					response.write "</article>"
					count2 = count2 + 1
				wend
				jigrs.close
				set jigrs = nothing
			next
			response.write "<count>" & count2 & "</count>"
		else
			Dim count : count = 0
			'---列出文章---
			jigsql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.iCUItem = KnowledgeJigsaw.gicuitem " & _
							 "WHERE (KnowledgeJigsaw.parentIcuitem = " & jigid & ") AND (KnowledgeJigsaw.Status = 'Y') AND (KnowledgeJigsaw.CtRootId = " & rid & ") " & _
							 "ORDER BY KnowledgeJigsaw.orderArticle DESC"
			set jigrs = conn.execute(jigsql)
			while not jigrs.eof 
				response.write "<article>"				
				response.write "<stitle>" & jigrs("stitle") & "</stitle>"				
				response.write "<path><![CDATA[" & jigrs("path") & "]]></path>"
				GetIssueArticlePostDate jigrs("ArticleId"), jigrs("CtRootId")
				jigrs.movenext
				response.write "</article>"
				count = count + 1
			wend
			jigrs.close
			set jigrs = nothing
			response.write "<count>" & count & "</count>"
		end if
		response.write "</" & atype & ">"
	end sub
	
	sub GetIssueArticlePostDate( id, rid )
		if rid = "4" then
			sql = "SELECT CONVERT(varchar, ONLINE_DATE, 111) AS ONLINE_DATE FROM REPORT WHERE REPORT_ID = '" & id & "'"
			set kmrs = kmconn.execute(sql)
			if not kmrs.eof then
				response.write "<xpostdate>" & kmrs("ONLINE_DATE") & "</xpostdate>"
			else 
				response.write "<xpostdate></xpostdate>"
			end if 
			kmrs.close
			set kmrs = nothing
		else
			sql = "SELECT CONVERT(varchar, xPostDate, 111) AS xPostDate FROM CuDTGeneric WHERE icuitem = " & id
			set newrs = conn.execute(sql)
			if not newrs.eof then
				response.write "<xpostdate>" & newrs("xPostDate") & "</xpostdate>"
			else
				response.write "<xpostdate></xpostdate>"
			end if
			newrs.close
			set newrs = nothing
		end if
		
	end sub
		
%>	
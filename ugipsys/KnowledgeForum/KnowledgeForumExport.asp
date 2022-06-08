<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	HTProgCap = "單元資料維護"
	HTProgFunc = "查詢"
	HTProgCode="kpi01"
	HTProgPrefix="" 
	response.charset = "UTF-8"
	Response.Buffer = False
%>
<!--#include virtual = "/inc/server.inc" -->
<%
	Dim filename : filename = replace( DateAdd("d", 0 , year(now) & "/" & month(now) & "/" & day(now) ), "/" , "")
	Response.AddHeader "Content-Disposition", "attachment;filename=" & filename & ".xls" 
	Response.ContentType = "application/vnd.ms-excel" 

	dim xNewWindow:xNewWindow=request("xNewWindow")
	dim sTitle:sTitle=Trim(request("sTitle"))
	dim memberId:memberId=request("memberId")
	dim Status:Status=request("Status")
	dim fCTUPublic:fCTUPublic=request("fCTUPublic")
	dim vGroup:vGroup=request("vGroup")
	dim xbody:xbody=Trim(request("xbody"))
	dim xPostDateS:xPostDateS=request("xPostDateS")
	dim xPostDateE:xPostDateE=request("xPostDateE")

	dim sTitleArray:sTitleArray = Split(sTitle," ")
	dim xBodyArray:xBodyArray = Split(xbody," ")
	Dim condition : condition = ""
	
	' if xNewWindow <> "" then response.write xNewWindow & "<br/>"
	' if sTitle <> "" then response.write sTitle &  "<br/>"
	' if memberId <> "" then response.write memberId &  "<br/>"
	' if Status <> "" then response.write Status & "<br/>"
	' if fCTUPublic <> "" then response.write fCTUPublic & "<br/>"
	' if vGroup <> "" then response.write vGroup & "<br/>"
	' if xbody <> "" then response.write xbody &  "<br/>"
	' if xPostDateS <> "" then response.write xPostDateS & "<br/>"
	' if xPostDateE <> "" then response.write xPostDateE & "<br/>"
	
	fsql= "SELECT DISTINCT CuDTGeneric.iCUItem, CuDTGeneric.xPostDate, CuDTGeneric.iBaseDSD, CuDTGeneric.iCTUnit, CuDTGeneric.fCTUPublic, CuDTGeneric.sTitle, CuDTGeneric.topCat, "
  fsql=fsql & "CuDTGeneric.iEditor, CuDTGeneric.xNewWindow, KnowledgeForum.gicuitem, KnowledgeForum.DiscussCount, KnowledgeForum.CommandCount,"
  fsql=fsql & "KnowledgeForum.BrowseCount, KnowledgeForum.TraceCount, KnowledgeForum.GradeCount, KnowledgeForum.ParentIcuitem, KnowledgeForum.Status,"
  fsql=fsql & "Member.account, Member.nickname, Member.realname, ISNULL(CuDTGeneric.vGroup,'') AS vGroup FROM  CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN"
  fsql=fsql & " Member ON CuDTGeneric.iEditor = Member.account WHERE (CuDTGeneric.iCTUnit = 932)"
  
  if xPostDateS <> "" then 
    fSql = fSql & " AND xPostDate >= '" & xPostDateS & " 00:00' " 
  end if

  if xPostDateE <> "" then 
    fSql = fSql & " AND xPostDate <= '" & xPostDateE & " 23:59' " 
  end if 

  if xPostDateS <> "" And xPostDateE <> "" then
	condition = condition & "發佈日期:" & xPostDateS & "~" & xPostDateE & ","
  elseif xPostDateS <> "" and xPostDateE = "" then
	condition = condition & "發佈日期:" & xPostDateS & "~至今,"
  elseif xPostDateS = "" and xPostDateE <> "" then
	condition = condition & "發佈日期:至" & xPostDateE & ","
  end if
  
  
  if xNewWindow <> "" then 
    fsql = fsql & " AND xNewWindow LIKE '%" & xNewWindow & "%' " 
	if xNewWindow = "N" then 
		condition = condition & "討論關閉:否," 
	else
		condition = condition & "討論關閉:是," 
	end if
  end if
  if sTitle <> "" then 
	if UBound(sTitleArray) > -1 then condition = condition & "問題標題:"  & sTitle & ","
    For i = LBound(sTitleArray) To UBound(sTitleArray)
    str = Right(sTitleArray(i),Len(sTitleArray(i)))
    if str <> "" then
      fSql = fSql & " AND CuDTGeneric.sTitle LIKE '%" & str & "%' "
	  
    end if
    Next
  
  end if
 

  if xbody <> "" then 
	if UBound(xBodyArray) > -1 then condition = condition & "內文:"  & xbody & ","
    For i = LBound(xBodyArray) To UBound(xBodyArray)
        str = trim(xBodyArray(i))
        if str <> "" then    
            fsql = fsql & " AND CuDTGeneric.xbody LIKE '%" & str & "%' "   
        end if
    Next  
  end if  
  
  
  if memberId <> "" then 
    fsql = fsql & " AND ( account LIKE '%" & memberId & "%' OR realname LIKE '%" & memberId & "%' "
    fsql = fsql & " OR nickname LIKE '%" & memberId & "%') "
	condition = condition & "上傳者:" & memberId & ","
  end if
  if Status <> "" then 
    fsql = fsql & " AND KnowledgeForum.Status LIKE '%" & Status & "%' " 
	if Status = "N" then 
		condition = condition & "狀態:正常," 
	else
		condition = condition & "狀態:刪除," 
	end if
  end if
  if fCTUPublic <> "" then 
    fsql = fsql & " AND fCTUPublic LIKE '%" & fCTUPublic & "%' " 
	if fCTUPublic = "N" then 
		condition = condition & "是否公開:否," 
	else
		condition = condition & "是否公開:是," 
	end if
  end if
  if vGroup <> "" then 
	if vGroup = "N" then
		fsql = fsql & " AND ISNULL(vGroup,'') = '' " 
		condition = condition & "活動問題狀態:否," 
	else
		fsql = fsql & " AND ISNULL(vGroup,'') LIKE '%" & vGroup & "%' "
		if vGroup = "A" then 
			condition = condition & "活動問題狀態:是," 
		else
			condition = condition & "活動問題狀態:已下架," 
		end if
	end if
    
  end if
  
  fsql=fsql & " ORDER BY CuDTGeneric.xPostDate DESC, CuDTGeneric.iCUItem DESC"
  
'response.write fsql
' response.write condition
	
	if len(condition) > 0 then condition = left(condition, len(condition) - 1)
	
	response.write "<table border=""1"">" & vbcrlf
	response.write "<tr><td colspan=""14""><font face=""新細明體"">匯出日期：" & Date() & "</font></td></tr>" & vbcrlf 
	response.write "<tr><td colspan=""14""><font face=""新細明體"">條件 => " & condition & "</font></td></tr>" & vbcrlf
	response.write "<tr><td colspan=""14""><font face=""新細明體"">&nbsp;</font></td>"
	response.write "<tr><td><font face=""新細明體"">&nbsp;</font></td>" & _
		        "<td><font face=""新細明體"">編號</font></td>" & _
						"<td><font face=""新細明體"">活動題目</font></td>" & _
						"<td><font face=""新細明體"">題目</font></td>" & _
						"<td><font face=""新細明體"">分類</font></td>" & _
						"<td><font face=""新細明體"">發佈時間</font></td>" & _
						"<td><font face=""新細明體"">討論數</font></td>" & _
						"<td colspan=""7""><font face=""新細明體"">內文</font></td></tr>" & vbcrlf
	
	set rs = conn.execute(fsql)	
	
	dim articleBody,vGroupCode,vGroupName,articletype,articletypeId,no
	no=0
	while not rs.eof
		no = no+1
		sql = "select xBody From CuDTGeneric where iCUItem=" & rs("iCUItem")
		set bodyRs = conn.execute(sql)	
		if not bodyRs.eof then articleBody = trim(bodyRs("xbody"))
		
		vGroupCode=""
		vGroupName=""
		if trim(rs("vGroup")) <> "" then
			vGroupCode = trim(rs("vGroup"))
			if vGroupCode = "A" then
				vGroupName="是"
			elseif vGroupCode = "OFF" then
				vGroupName="已下架"
			end if
		else
			vGroupName="否"
		end if
		
		articletypeId=""
		articletype=""
		if trim(rs("topCat")) <> "" then
			articletypeId = trim(rs("topCat"))
			if articletypeId = "A" then
				articletype="農"
			elseif articletypeId = "B" then
				articletype="林"
			elseif articletypeId = "C" then
				articletype="漁"
			elseif articletypeId = "D" then
				articletype="牧"
			elseif articletypeId = "E" then
				articletype="其他"
			end if
		end if
		
	
		response.write "<tr><td><font face=""新細明體"">&nbsp;" & no & "</font></td>" & _
                   "<td><font face=""新細明體"">&nbsp;" & rs("iCUItem") & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & vGroupName & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & trim(rs("sTitle")) & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & articletype & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & trim(rs("xPostDate")) & "</font></td>" & _
									 "<td><font face=""新細明體"">" & CInt(rs("DiscussCount")) & "</font></td>" & _
									 "<td colspan=""7""><font face=""新細明體"">&nbsp;" & articleBody & "</font></td></tr>" & vbcrlf
		rs.movenext
	wend
	rs.close
	set rs = nothing
	
	response.write "</table>" & vbcrlf
	
	
%>

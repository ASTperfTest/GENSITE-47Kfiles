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
<!--#include virtual = "/inc/client.inc" -->
<%
	Dim filename : filename = replace( DateAdd("d", 0 , year(now) & "/" & month(now) & "/" & day(now) ), "/" , "")
	Response.AddHeader "Content-Disposition", "attachment;filename=" & filename & ".xls" 
	Response.ContentType = "application/vnd.ms-excel" 

	dim memberId:memberId=request("MemberId")
	dim memberIdEncode:memberIdEncode=request("MemberIdEncode")
	dim scoreS:scoreS=request("ScoreS")
	dim scoreE:scoreE=request("ScoreE")
	
'response.write Chg_UNI(memberId)
'response.write "'" & Chg_UNI(memberId) & "'"
'response.write ascw(mid(memberId,1,1)) & "<br/>"
'response.write "&#" & ascw(mid(memberId,1,1)) & ";"  & "<br/>"

	Dim condition : condition = ""
	
	fsql = "SELECT Total.MemberId,M.nickname,M.realname,SUM(Total.Grade) As Grade,M.email,SUM(Total.QuestionNum) AS QuestionNum,SUM(Total.DiscussNum) AS DiscussNum "
	fsql=fsql &" FROM (	SELECT KA.MemberId,COUNT(KA.MemberId) AS  QuestionNum,'' AS DiscussNum,SUM(KA.Grade) AS Grade FROM dbo.KnowledgeActivity KA "
	fsql=fsql & "WHERE type BETWEEN 1 AND 2 "
	fsql=fsql & "GROUP BY KA.MemberId "
	fsql=fsql & "UNION ALL "
	fsql=fsql & "SELECT Temp.MemberId,'' AS  QuestionNum,COUNT(Temp.MemberId)AS DiscussNum,SUM(temp.Grade) AS Grade "
	fsql=fsql & "FROM( SELECT KA.MemberId,CASE  WHEN SUM(KA.Grade)>4 THEN 4 ELSE SUM(KA.Grade) END AS Grade  "
	fsql=fsql & "FROM dbo.KnowledgeActivity KA "
	fsql=fsql & "INNER JOIN dbo.KnowledgeForum KF ON ka.CUItemId = KF.gicuitem "
	fsql=fsql & "WHERE KA.State =1 AND KA.type BETWEEN 3 AND 4 AND KF.Status = 'N' "
	fsql=fsql & "GROUP BY KF.ParentIcuitem, KA.MemberId) AS Temp "
	fsql=fsql & "GROUP BY Temp.MemberId) Total "
	fsql=fsql & "INNER JOIN dbo.Member M ON Total.MemberId = M.account "
	fsql=fsql & "WHERE M.status <> 'N'"
	
	
  if memberId <> "" then
	condition = condition & "會員:" & memberId & ","
    fSql = fSql & "AND ( Total.MemberId LIKE '%" & memberId & "%' OR (M.realname LIKE '%" & memberId & "%' OR M.realname LIKE '%" & Chg_UNI(memberId) & "%') OR (M.nickname LIKE '%" & memberId & "%' OR M.nickname LIKE '%" & Chg_UNI(memberId) & "%')) "
  end if
  
	fSql = fSql & "GROUP BY Total.MemberId,M.nickname,M.realname,M.email " 
  
  if scoreS <> "" AND scoreE <> "" then 
	condition = condition & "得分:" & scoreS & "~" & scoreE  & ","
   fSql = fSql &  "HAVING SUM(Total.Grade) Between " & scoreS & " AND " & scoreE 
  end if
  
  
  fSql = fSql & " ORDER BY Grade desc"
  
  
'response.write fsql
' response.write condition
	
	if len(condition) > 0 then condition = left(condition, len(condition) - 1)
	
	response.write "<table border=""1"">" & vbcrlf
	response.write "<tr><td colspan=""8""><font face=""新細明體"">匯出日期：" & Date() & "</font></td></tr>" & vbcrlf 
	response.write "<tr><td colspan=""8""><font face=""新細明體"">條件 => " & condition & "</font></td></tr>" & vbcrlf
	response.write "<tr><td colspan=""8""><font face=""新細明體"">&nbsp;</font></td>"
	response.write "<tr><td><font face=""新細明體"">&nbsp;</font></td>" & _
						"<td><font face=""新細明體"">帳號</font></td>" & _
						"<td><font face=""新細明體"">暱稱</font></td>" & _
						"<td><font face=""新細明體"">姓名</font></td>" & _
						"<td><font face=""新細明體"">E-mail</font></td>" & _
						"<td><font face=""新細明體"">發問數</font></td>" & _
						"<td><font face=""新細明體"">討論數</font></td>" & _
						"<td><font face=""新細明體"">得分</font></td></tr>" & vbcrlf
						
	
	
	set rs = conn.execute(fsql)	
	
	dim strAcount,strNickname,strRealname,no
	 no=0
	
	
	 while not rs.eof
	 no = no+1
	 response.write "<tr><td><font face=""新細明體"">&nbsp;" & no & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & rs("MemberId") & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & rs("nickname") & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & rs("realname") & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & rs("email") & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & rs("QuestionNum") & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & rs("DiscussNum") & "</font></td>" & _
									 "<td><font face=""新細明體"">&nbsp;" & trim(rs("Grade")) & "</font></td></tr>" & vbcrlf
		rs.movenext
	wend
	rs.close
	set rs = nothing
	
	response.write "</table>" & vbcrlf
	
	
Function Chg_UNI(str)        'ASCII轉Unicode
	dim old,new_w,iStr
	old = str
	new_w = ""
	
	for iStr = 1 to len(str)
		'response.write ascw(mid(old,iStr,1)) & "<br/>"
		if ascw(mid(old,iStr,1)) < 0 then
			'response.write "1" & "<br/>"
			new_w = new_w & "&#" & ascw(mid(old,iStr,1))+65536 & ";"
		elseif        ascw(mid(old,iStr,1))>0 and ascw(mid(old,iStr,1))<127 then
			'response.write "2" & "<br/>"
			new_w = new_w & mid(old,iStr,1)
		else
			'response.write "3" & "<br/>"
			new_w = new_w & "&#" & ascw(mid(old,iStr,1)) & ";"
			'response.write new_w & "<br/>"
		end if
	next
	
	'response.write new_w & "<br/>"
	Chg_UNI=new_w
End Function
	
%>

<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
HTProgCap=""
HTProgFunc=""
HTProgCode="KForumlist"
HTProgPrefix=""
HTUploadPath=session("Public")+"data/"
%>
<!--#include virtual = "/inc/server.inc" -->
<!-- #INCLUDE FILE="KpiFunction.inc" -->
<% 
	
function xUpForm(xvar)
	xUpForm = trim(request(xvar))
end function

dim iCUItem
iCUItem = request.querystring("iCUItem")	
		
if request("submitTask") = "delete" then
	
	sql = "SELECT Status FROM KnowledgeForum WHERE gicuitem = " & iCUItem	
	set rs = conn.execute(sql)
	if not rs.eof then
		if rs("Status") = "D" then 
			response.write "<script>alert('article can not be deleted!!');window.location.href='KnowledgeForumlist.asp';</script>"
			response.end
		end if
	end if
	'---kpi---
	'------檢查是否有討論------
	sql = "SELECT CuDTGeneric.iCUItem FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
	sql = sql & "WHERE (KnowledgeForum.ParentIcuitem = " & iCUItem & ") AND (KnowledgeForum.Status = 'N') AND (CuDTGeneric.iCtUnit = 933)"
	set rs = conn.execute(sql)
	while not rs.eof 
		DeleteCommend rs("iCUItem") '---刪除評價---	
		DeleteOpinion rs("iCUItem") '---刪除意見---			
		DeleteDiscuss rs("iCUItem") '---刪除討論---		
		rs.movenext
	wend
	rs.close
	set rs = nothing
	'-----------------------------------------------------------------------------------		
	'---刪除發問---
	DeleteQuestion iCUItem	
	UpdateStatus iCUItem
	'-----------------------------------------------------------------------------------
	'---for 知識家的刪除-刪題目或是討論或是意見---
	'---do nothing---因為題目本身已看不到---	
	'---end of for 知識家的刪除---
	'-----------------------------------------------------------------------------------
	'---資料刪除時,要將資料的index刪除...先將這筆資料加到DB中---
	
	'-----------------------------------------------------------------------------------	
	showDoneBox "刪除成功", "KnowledgeForumlist.asp"
end if

if request("submitTask") = "UPDATE" then
	
	stitle=replace(xUpForm("stitle"),"'","''")
	xbody=replace(xUpForm("htx_xbody"),"'","''")
	xurl=replace(xUpForm("htx_xurl"),"'","''")
	xnewWindow=replace(xUpForm("htx_xnewWindow"),"'","''")
	fctupublic=replace(xUpForm("htx_fctupublic"),"'","''")
	topCat=replace(xUpForm("htx_topCat"),"'","''")
	vgroup=replace(xUpForm("htx_vgroup"),"'","''")
	xkeyword=replace(xUpForm("htx_xkeyword"),"'","''")
	ximportant=replace(xUpForm("htx_ximportant"),"'","''")
	idept=replace(xUpForm("htx_idept"),"'","''")

	sql = "UPDATE CuDTGeneric SET  "
	sql = sql  & "stitle =" & "' " & stitle &"'"
	sql = sql  & ",xbody =" & "' " & xbody &"'"
	sql = sql  & ",xurl = "& "'" & xurl  &"'"
	sql = sql  & ",xnewWindow = "& "'" & xnewWindow  &"'"
	sql = sql  & ",fctupublic = "&"'" & fctupublic  &"'"
	sql = sql  & ",topCat = "&"'" & topCat  &"'"
	sql = sql  & ",vgroup = "&"'" & vgroup  &"'"
	sql = sql  &",xkeyword = "& "'" & xkeyword  &"'"
	sql = sql  &",ximportant = "& "'" & ximportant  &"'"
	sql = sql  &",idept = "& "'" & idept  &"'"
	sql = sql  & " WHERE iCUItem = "& "'"  & iCUItem & "'"
	conn.execute(sql)
	showDoneBox "更新成功", "KnowledgeForumlist.asp"	
end if

%>
<% Sub showDoneBox(lMsg, link) %>
  <html>
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<meta name="GENERATOR" content="Hometown Code Generator 1.0">			
			<title>編修表單</title>
    </head>
    <body>
			<script language=vbs>
			  alert("<%=lMsg%>")	    				
				document.location.href = link
			</script>
    </body>
  </html>
<% End sub %>
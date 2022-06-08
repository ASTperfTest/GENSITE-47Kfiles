<%
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCode="KForumlist"
HTProgPrefix =" KnowledgeForum"
HTUploadPath=session("Public")+"data/"
%>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../css/list.css" rel="stylesheet" type="text/css">
<link href="../css/layout.css" rel="stylesheet" type="text/css">
<title>資料管理／資料上稿</title>
</head>
<% Sub showCantFindBox(lMsg) %>
  <html>
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<meta name="GENERATOR" content="Hometown Code Generator 1.0">
			<meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
			<title>編修表單</title>
    </head>
    <body>
			<script language=vbs>
			  alert("<%=lMsg%>")	    				
				history.back
			</script>
    </body>
  </html>
<% End sub %>
<%

dim picId:picId=request("picId")

dim xNewWindow:xNewWindow=request("xNewWindow")
dim validate:validate=request("validate")
dim sTitle:sTitle=request("sTitle") 
dim memberId:memberId=request("memberId")
dim Status:Status=request("Status")
dim fCTUPublic:fCTUPublic=request("fCTUPublic")
dim xbody:xbody=request("xbody")
	
'cSql = "SELECT COUNT(*) as allcount FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN KnowledgePicInfo ON KnowledgeForum.gicuitem = KnowledgePicInfo.parentIcuitem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account WHERE (CuDTGeneric.iCTUnit = 932) "
cSql = "SELECT COUNT(*) as allcount FROM  CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account WHERE (CuDTGeneric.iCTUnit = 932)"
	
 '---加入搜尋條件-
  'if validate <> ""   then 
     'fSql = fSql & " AND picStatus LIKE '%" & validate & "%' " 
	 'cSql = cSql & " AND picStatus LIKE '%" & validate & "%' " 
  'end if

  if xNewWindow <> "" then 
    fSql = fSql & " AND xNewWindow LIKE '%" & xNewWindow & "%' " 
    cSql = cSql & " AND xNewWindow LIKE '%" & xNewWindow & "%' " 
  end if
  if sTitle <> "" then 
    fSql = fSql & " AND CuDTGeneric.sTitle LIKE '%" & sTitle & "%' " 
    cSql = cSql & " AND CuDTGeneric.sTitle LIKE '%" & sTitle & "%' " 
  end if
  if xbody <> "" then 
    fSql = fSql & " AND CuDTGeneric.xbody LIKE '%" & xbody & "%' " 
    cSql = cSql & " AND CuDTGeneric.xbody LIKE '%" & xbody & "%' " 
  end if
  if memberId <> "" then 
    fSql = fSql & " AND ( account LIKE '%" & memberId & "%' OR realname LIKE '%" & memberId & "%' "
    fSql = fSql & " OR nickname LIKE '%" & memberId & "%') "
    cSql = cSql & " AND ( account LIKE '%" & memberId & "%' OR realname LIKE '%" & memberId & "%' "
    cSql = cSql & " OR nickname LIKE '%" & memberId & "%') "
  end if
  if Status <> "" then 
    fSql = fSql & " AND KnowledgeForum.Status LIKE '%" & Status & "%' " 
    cSql = cSql & " AND KnowledgeForum.Status LIKE '%" & Status & "%' " 
  end if
  
  if fCTUPublic <> "" then 
    fSql = fSql & " AND fCTUPublic LIKE '%" & fCTUPublic & "%' " 
    cSql = cSql & " AND fCTUPublic LIKE '%" & fCTUPublic & "%' " 
  end if

  nowPage = Request.QueryString("nowPage")  '現在頁數
  PerPageSize = cint(Request.QueryString("pagesize"))
  if PerPageSize <= 0 then  PerPageSize = 15  

  set RSc = conn.execute(cSql)
  totRec = RSc(0)       '總筆數
  totPage = int(totRec / PerPageSize + 0.999)

  if cint(nowPage) < 1 then 
    nowPage = 1
  elseif cint(nowPage) > totPage then 
    nowPage = totPage 
  end if      
  
  ''fSql= "SELECT TOP "& nowPage * PerPageSize  & " CuDTGeneric.iCUItem, CuDTGeneric.iBaseDSD, CuDTGeneric.iCTUnit, CuDTGeneric.fCTUPublic, CuDTGeneric.sTitle, CuDTGeneric.iEditor,CuDTGeneric.xNewWindow, KnowledgeForum.gicuitem, KnowledgeForum.DiscussCount, KnowledgeForum.CommandCount,KnowledgeForum.BrowseCount, KnowledgeForum.TraceCount, KnowledgeForum.GradeCount, KnowledgeForum.ParentIcuitem, KnowledgeForum.Status,KnowledgePicInfo.picId, KnowledgePicInfo.picTitle, KnowledgePicInfo.parentIcuitem , Member.account, Member.nickname, Member.realname,KnowledgePicInfo.picStatus FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN KnowledgePicInfo ON KnowledgeForum.gicuitem = KnowledgePicInfo.parentIcuitem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account WHERE (CuDTGeneric.iCTUnit = 932)" & fSql & " ORDER BY CuDTGeneric.xPostDate DESC, CuDTGeneric.iCUItem DESC"
	csql=csql&" GROUP BY CuDTGeneric.iCTUnit" 
	fsql= "SELECT DISTINCT TOP "& nowPage * PerPageSize  & " CuDTGeneric.iCUItem, CuDTGeneric.xPostDate, CuDTGeneric.iBaseDSD, CuDTGeneric.iCTUnit, CuDTGeneric.fCTUPublic, CuDTGeneric.sTitle, CuDTGeneric.topCat, "
  fsql=fsql & "CuDTGeneric.iEditor, CuDTGeneric.xNewWindow, KnowledgeForum.gicuitem, KnowledgeForum.DiscussCount, KnowledgeForum.CommandCount,"
  fsql=fsql & "KnowledgeForum.BrowseCount, KnowledgeForum.TraceCount, KnowledgeForum.GradeCount, KnowledgeForum.ParentIcuitem, KnowledgeForum.Status,"
  fsql=fsql & "Member.account, Member.nickname, Member.realname FROM  CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN"
  fsql=fsql & " Member ON CuDTGeneric.iEditor = Member.account WHERE (CuDTGeneric.iCTUnit = 932)"
	   
if xNewWindow <> "" then 
    fsql = fsql & " AND xNewWindow LIKE '%" & xNewWindow & "%' " 
  end if
  if sTitle <> "" then 
    fsql = fsql & " AND CuDTGeneric.sTitle LIKE '%" & sTitle & "%' " 
  end if
  if xbody <> "" then 
    fsql = fsql & " AND CuDTGeneric.xbody LIKE '%" & xbody & "%' " 
  end if
  if memberId <> "" then 
    fsql = fsql & " AND ( account LIKE '%" & memberId & "%' OR realname LIKE '%" & memberId & "%' "
    fsql = fsql & " OR nickname LIKE '%" & memberId & "%') "
  end if
  if Status <> "" then 
    fsql = fsql & " AND KnowledgeForum.Status LIKE '%" & Status & "%' " 
  end if
  if fCTUPublic <> "" then 
    fsql = fsql & " AND fCTUPublic LIKE '%" & fCTUPublic & "%' " 
  end if
  fsql=fsql & " ORDER BY CuDTGeneric.xPostDate DESC, CuDTGeneric.iCUItem DESC"
  
	'response.write fsql &"<hr>"
	'response.write csql
	'response.end

  Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
'	RSreg.CacheSize = PerPageSize

'----------HyWeb GIP DB CONNECTION PATCH----------
'	RSreg.Open fsql,Conn, 3, 1
Set RSreg = Conn.execute(fsql)
	RSreg.CacheSize = PerPageSize
'----------HyWeb GIP DB CONNECTION PATCH----------


	if Not RSreg.eof then
		if totRec > 0 then 
      RSreg.PageSize = PerPageSize       '每頁筆數
      RSreg.AbsolutePage = nowPage      
		end if    
	end if 
	
	if RSreg.eof then
		showCantFindBox "找不到資料"    
	else
%>
<body>

<div id="FuncName">
	<h1>資料管理／資料上稿</h1>
	<font size=2>【目錄樹節點: 知識問題】
	<div id="Nav">
			<A href="knowledgeQueryTopic.asp" title="指定查詢條件">查詢</A>			
	<div id="ClearFloat"></div>
</div>

<div id="FormName">
	單元資料維護&nbsp;
	<font size=2>【主題單元:(2008)知識問題 / 單元資料:純網頁】
</div>
<!--  條列頁簡易查詢功能  -->


<Form id="Form2" name=reg method="POST" action="KnowledgeForumList.asp?keep=Y&nowPage=<%=(nowPage)%>&pagesize=<%=PerPageSize%>&xNewWindow=<%=trim(request("xNewWindow"))%>&sTitle=<%=Server.URLEncode(trim(request("sTitle")))%>&memberId=<%=Server.URLEncode(memberId)%>&fCTUPublic=<%=request("fCTUPublic")%>&Status=<%=request("Status")%>&xbody=<%=Server.URLEncode(request("xbody"))%>">
	<INPUT TYPE=hidden name=submitTask value="">
	<!-- 分頁 -->
	<div id="Page">

	共<em><%=totRec%></em>筆資料，目前在第<em><%=nowPage%> /<%=totPage%></em>	頁	
	每頁顯示  
	<select id=PerPage size="1" style="color:#FF0000" class="select">            
    <option value="15"<%if PerPageSize=15 then%> selected<%end if%>>15</option>                       
    <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
    <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
	</select> 
	筆
	<% if cint(nowPage) <> 1 then %>
		<img src="/images/arrow_previous.gif" alt="上一頁">       		
		<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>&xNewWindow=<%=trim(request("xNewWindow"))%>&sTitle=<%=Server.URLEncode(trim(request("sTitle")))%>&memberId=<%=Server.URLEncode(memberId)%>&fCTUPublic=<%=request("fCTUPublic")%>&Status=<%=request("Status")%>&xbody=<%=Server.URLEncode(request("xbody"))%>">上一頁</a> ｜
  <% end if %>跳至第
  <select id=GoPage size="1" style="color:#FF0000" class="select">
	<% For iPage=1 to totPage %> 
		<option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
  <% Next %>   
  </select>      
	頁	
	<% if cint(nowPage)<>totPage then %> 
   ｜<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>&xNewWindow=<%=trim(request("xNewWindow"))%>&sTitle=<%=Server.URLEncode(trim(request("sTitle")))%>&memberId=<%=Server.URLEncode(memberId)%>&fCTUPublic=<%=request("fCTUPublic")%>&Status=<%=request("Status")%>&xbody=<%=Server.URLEncode(request("xbody"))%>">下一頁
    <img src="/images/arrow_next.gif" alt="下一頁"></a> 
  <% end if %>  
		
  </div>
	<!-- 列表 -->
<table cellspacing="0" id="ListTable">
<tr>
	<th class="First" scope="col">預覽</th>
	<th><p align="center">標 題</p></th>
	<th><p align="center">討論數</th>
	<th><p align="center">瀏覽數</th>
	<th><p align="center">狀態</th>
	<th><p align="center">討論關閉</th>
	<th><p align="center">是否公開(草稿)</th>
	<th><p align="center">圖片狀態</th>
	<th><p align="center">發表帳號</th>
</tr>                  
<tr>         
<% 
	while not RSreg.eof 
		weburl = session("myWWWSiteURL") & "/knowledge/knowledge_cp.aspx?ArticleId={0}&ArticleType=A&CategoryId={1}&kpi=0"
		weburl = replace( weburl, "{0}", RSreg("iCUItem") )
		weburl = replace( weburl, "{1}", RSreg("topCat") )
%>
  <td align="center"><A href="<%=weburl%>" target="_wMof">View</A></td> 
	<td align="center"><a href="cp_question.asp?iCUItem=<%=RSreg("iCUItem")%>&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>"><%=RSreg("sTitle")%></a></td>
	<td align="center"><%=RSreg("DiscussCount")%></td>
	<td align="center"><%=RSreg("BrowseCount")%></td>
	<td align="center">
		<% if RSreg("status")="N" then %>正常<% end if %>
		<% if RSreg("status")="D" then %>刪除<% end if %>
	</td>
	<td align="center">
		<% if RSreg("xNewWindow")="Y" then%>是<% END IF %>
		<% if RSreg("xNewWindow")="N" then%>否<% END IF %>	
	</td>
  <td align="center">
		<% if RSreg("fCTUPublic")="Y" then%>公開<% END IF %>
		<% if RSreg("fCTUPublic")="N" then%>不公開<% END IF %>	
	</td>	
<%

sql="SELECT KnowledgeForum.gicuitem, KnowledgePicInfo.picStatus FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN KnowledgePicInfo ON KnowledgeForum.gicuitem = KnowledgePicInfo.parentIcuitem "
sql =sql & "WHERE  (CuDTGeneric.iCTUnit = 932) and CuDTGeneric.iCUItem = "& "'"  & RSreg("iCUItem") & "'"

Set RS = conn.execute(sql)
FLAG=0

if not rs.eof then

	while not RS.eof
		if RS("picStatus")= "W" then
			FLAG=1
		'elseif trim(RS("picStatus")) = "" then
			'FLAG=2
		end if
		RS.MoveNext
	wend
else
	flag = 2
end if

response.write "<td align=""center"">"
if FLAG=0 then
response.write "已審核"
elseif (FLAG=1) then
response.write "待審核"
elseif (FLAG=2) then
response.write "無圖"
END IF 
response.write "</td>"

%> 
   
	<td align="center"><%=RSreg("iEditor")%></td>
	
	
</tr>
 <%	
			RSreg.MoveNext
		wend
 %>                    
           
	
</tr>
    </TABLE>

<!-- 程式結束 ---------------------------------->  
</form>  
</body>
</html>      

<script Language=VBScript>
	
		sub formModSubmit()
			reg.submitTask.value = "DELETE"
			reg.action = "<%=HTProgPrefix%>List.asp?icuitem=<%=request.querystring("icuitem")%>"
			reg.Submit
		end sub
		
		dim gpKey
    sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=" & newPage & "&pagesize=<%=PerPageSize%>" _
           &"&xNewWindow=<%=trim(request("xNewWindow"))%>&sTitle=<%=Server.URLEncode(trim(request("sTitle")))%>&memberId=<%=Server.URLEncode(memberId)%>&fCTUPublic=<%=request("fCTUPublic")%>&Status=<%=request("Status")%>"		   
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>" & "&pagesize=" & newPerPage _
           &"&xNewWindow=<%=trim(request("xNewWindow"))%>&sTitle=<%=Server.URLEncode(trim(request("sTitle")))%>&memberId=<%=Server.URLEncode(memberId)%>&fCTUPublic=<%=request("fCTUPublic")%>&Status=<%=request("Status")%>"		   
     end sub 

     sub shortLongList(sorl)
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>" _
			& "&shortLongList=" & sorl
     end sub     	
	</script>                                       
<% end if %>                           

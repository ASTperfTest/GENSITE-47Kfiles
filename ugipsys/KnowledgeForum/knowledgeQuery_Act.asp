<% 
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTUploadPath=session("Public")+"data/"
HTProgPrefix = "knowledgeQuery_Act"
Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="KForumlist"

%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/client.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../css/list.css" rel="stylesheet" type="text/css">
<link href="../css/layout.css" rel="stylesheet" type="text/css">
<link href="../css/form.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>
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

'---計算筆數-
cSql = "SELECT COUNT(*) AS allcount FROM Member INNER JOIN CuDTGeneric ON Member.account = CuDTGeneric.iEditor " & _
"INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem " & _
"INNER JOIN KnowledgePicInfo ON KnowledgeForum.gicuitem = KnowledgePicInfo.parentIcuitem " & _
"WHERE ((CuDTGeneric.iCTUnit = 932) OR (CuDTGeneric.iCTUnit = 933) OR (CuDTGeneric.iCTUnit = 934))"

fsql=" ISNULL(KnowledgeForum.parentIcuitem,'0') ArticleParentIcuitem,CuDTGeneric.iCUItem,CuDTGeneric.sTitle, KnowledgePicInfo.picStatus, Member.account, Member.realname, Member.nickname, " & _
"KnowledgePicInfo.picTitle, KnowledgePicInfo.picId,KnowledgePicInfo.parentIcuitem,CuDTGeneric.topCat,KnowledgePicInfo.picPath " & _
"FROM Member INNER JOIN CuDTGeneric ON Member.account = CuDTGeneric.iEditor " & _
"INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem " & _
"INNER JOIN KnowledgePicInfo ON KnowledgeForum.gicuitem = KnowledgePicInfo.parentIcuitem " & _
"WHERE ((CuDTGeneric.iCTUnit = 932) OR (CuDTGeneric.iCTUnit = 933) OR (CuDTGeneric.iCTUnit = 934))"
dim picId:picId=request("picId")
dim PicidTerms:PicidTerms=request("PicidTerms")
dim parentIcuitem:parentIcuitem=request("parentIcuitem")
dim validate:validate=request("validate")
dim picTitle:picTitle=request("picTitle") 
dim memberId:memberId=request("memberId")
' response.write request("picId") & "<br/>"
' response.write request("PicidTerms") & "<br/>"
' response.write request("parentIcuitem") & "<br/>"
' response.write request("validate") & "<br/>"
' response.write request("picTitle") & "<br/>"
' response.write request("memberId") & "<br/>"
  
		'---更新圖片狀態---
if request("submitTask") = "Pass" then
	if cint(request("totalRecord"))	= 1 then 
		flag="noQuery" 'totalRecord若為一筆,當該筆資料審核過後,即不再需要加入查詢條件
	else
		flag="1"
	end if
	sql = "UPDATE KnowledgePicInfo SET picStatus = 'Y' WHERE parentIcuitem = "& "'"  & parentIcuitem & "'" & " and picId= "& picId 
	conn.execute(sql)
elseif request("submitTask") = "NoPass" then
	if cint(request("totalRecord"))	= 1 then 
		flag="noQuery" 'totalRecord若為一筆,當該筆資料審核過後,即不再需要加入查詢條件
	else
		flag="1"
	end if
    sql = "UPDATE KnowledgePicInfo SET picStatus = 'N' WHERE parentIcuitem = "& "'"  & parentIcuitem & "'" & " and picId= "& picId 
	conn.execute(sql)
elseif request("submitTask") = "AllPass" then
	
	Dim selectedItems : selectedItems = request("selectedItems")
	Dim items : items = split(selectedItems, ";")
	for each item in items
		if item <> "" then 
			sql = "UPDATE KnowledgePicInfo SET picStatus = 'Y' WHERE picId=  " & item
			conn.execute(sql)
			
		end if
	next
elseif request("submitTask") = "AllNoPass" then
	
	Dim selectedItem : selectedItem = request("selectedItems")
	Dim Item : Items = split(selectedItem, ";")
	for each item in Items
		if Item <> "" then 
			sql = "UPDATE KnowledgePicInfo SET picStatus = 'N' WHERE picId=  " & Item
			conn.execute(sql)
			
		end if
	next
end if

	 '---加入搜尋條件-
if flag="1" then 
  if validate <> ""   then 
     fSql = fSql & " AND picStatus LIKE '%" & validate & "%' " 
	 cSql = cSql & " AND picStatus LIKE '%" & validate & "%' " 
  end if
   if picTitle <> "" then 
     fSql = fSql & " AND picTitle LIKE '%" & picTitle & "%' " 
     cSql = cSql & " AND picTitle LIKE '%" & picTitle & "%' " 
   end if
   if memberId <> "" then 
   fSql = fSql & " AND ( account LIKE '%" & memberId & "%' OR realname LIKE '%" & memberId & "%' "
   fSql = fSql & " OR nickname LIKE '%" & memberId & "%') "
   cSql = cSql & " AND ( account LIKE '%" & memberId & "%' OR realname LIKE '%" & memberId & "%' "
   cSql = cSql & " OR nickname LIKE '%" & memberId & "%') "
  end if
    if PicidTerms <> "" then
     fSql = fSql & " AND picid LIKE '%" & PicidTerms & "%' " 
	 cSql = cSql & " AND picid LIKE '%" & PicidTerms & "%' " 
  end if
  
elseif flag="noQuery" then
'do Nothing
else

  if validate <> ""   then 
     fSql = fSql & " AND picStatus LIKE '%" & validate & "%' " 
	 cSql = cSql & " AND picStatus LIKE '%" & validate & "%' " 
  end if
  if PicidTerms <> "" then
     fSql = fSql & " AND picid LIKE '%" & PicidTerms & "%' " 
	 cSql = cSql & " AND picid LIKE '%" & PicidTerms & "%' " 
  end if
  if picTitle <> "" then 
     fSql = fSql & " AND picTitle LIKE '%" & picTitle & "%' " 
     cSql = cSql & " AND picTitle LIKE '%" & picTitle & "%' " 
  end if
  if memberId <> "" then 
   fSql = fSql & " AND ( account LIKE '%" & memberId & "%' OR realname LIKE '%" & memberId & "%' "
   fSql = fSql & " OR nickname LIKE '%" & memberId & "%') "
   cSql = cSql & " AND ( account LIKE '%" & memberId & "%' OR realname LIKE '%" & memberId & "%' "
   cSql = cSql & " OR nickname LIKE '%" & memberId & "%') "
  end if

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
  
  fsql = "SELECT TOP "& nowPage * PerPageSize & fsql & " ORDER BY xPostDate DESC"


  Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
'	RSreg.CacheSize = PerPageSize

'----------HyWeb GIP DB CONNECTION PATCH----------
'	RSreg.Open fSql,Conn, 3, 1
Set RSreg = Conn.execute(fSql)
	RSreg.CacheSize = PerPageSize
'----------HyWeb GIP DB CONNECTION PATCH----------


	if Not RSreg.eof then
		if totRec > 0 then 
      RSreg.PageSize = PerPageSize       '每頁筆數
      RSreg.AbsolutePage = nowPage      
		end if    
	end if 
	
	'response.write fSql
if  RSreg.eof then
	showCantFindBox "找不到資料"      
else
	
%>
<body>
<div id="FuncName">
	<h1>圖片整合管理</h1>
	<div id="Nav">
			<A href="KnowledgeQuery.asp" title="指定查詢條件">圖片查詢</A>
			<!--A href="cp_member_new.htm" title="新增資料">新增會員</A-->
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">
	圖片管理&nbsp;
	<font size="2">【主題單元:知識圖片 / 單元資料:純網頁】</font></div>
<!--  條列頁簡易查詢功能  -->


<Form id="reg" name="reg" method="POST" action="knowledgeQuery_Act.asp?keep=Y&nowPage=<%=(nowPage)%>&pagesize=<%=PerPageSize%>&validate=<%=request("validate")%>&picTitle=<%=Server.URLEncode(picTitle)%>&memberId=<%=Server.URLEncode(memberId)%>&PicidTerms=<%=request("PicidTerms")%>">
 <INPUT TYPE=hidden id="submitTask" name="submitTask" value="">
  <INPUT TYPE=hidden id="totalRecord" name="totalRecord" value="">
 <INPUT TYPE=hidden  id="picId" name="picId" value="">
  <INPUT TYPE=hidden  id="parentIcuitem" name="parentIcuitem" value="">
  <INPUT TYPE=hidden id="selectedItems" name="selectedItems" value="">

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
			<a href="<%=HTProgPrefix%>.asp?keep=Y&nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>&validate=<%=trim(request("validate"))%>&picTitle=<%=Server.URLEncode(trim(request("picTitle")))%>&memberId=<%=Server.URLEncode(memberId)%>&PicidTerms=<%=request("PicidTerms")%>">上一頁</a> ｜
    <% end if %>跳至第
    <select id=GoPage size="1" style="color:#FF0000" class="select">
		<% For iPage=1 to totPage %> 
			<option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
    <% Next %>   
    </select>      
		頁	
	<% if cint(nowPage)<>totPage then %> 
     ｜<a href="<%=HTProgPrefix%>.asp?keep=Y&nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>&validate=<%=trim(request("validate"))%>&picTitle=<%=Server.URLEncode(trim(request("picTitle")))%>&memberId=<%=Server.URLEncode(memberId)%>&PicidTerms=<%=request("PicidTerms")%>">下一頁
      <img src="/images/arrow_next.gif" alt="下一頁"></a> 
    <% end if %>  
		
  </div>

  
	<!-- 列表 -->
	<table cellspacing="0" id="ListTable">
		<tr>

		<th class="First" scope="col">&nbsp;</th>
					<th scope="col">上傳帳號</th>
					<th scope="col">真實姓名</th>
					<th scope="col">會員暱稱</th>
					<th class=eTableLable>知識家問題</th>
					<th nowrap scope="col">圖片</th>
					<th nowrap scope="col">圖片上傳資料ID</th>
					<th nowrap scope="col">審核</th>
					<th nowrap scope="col">狀態</th>
					<!--<th class="First" scope="col">預覽</th>   -->
	  </tr>     

<% while not RsREG.eof 
	dim articleId
	if trim(RSreg("ArticleParentIcuitem")) <> "0" then
		articleId = trim(RSreg("ArticleParentIcuitem"))
	else
		articleId = trim(RSreg("iCUItem"))
	end if
%>	  
<tr>
  <td align="center"><span class="eTableContent">
    <input type="checkbox" name="checkbox"  id="<%=RSreg("picId")%>" value="checkbox">
  </span></td>                  
	   
	<td align="center"><%=trim(RSreg("account"))%></td>
	<td align="center"><%=trim(RSreg("realname"))%></td>
	<td align="center"><%=trim(RSreg("nickname"))%></td>
	<TD class=eTableContent><a href="http://kmweb.coa.gov.tw/knowledge/knowledge_cp.aspx?ArticleId=<%=articleId%>&ArticleType=A&CategoryId=A&kpi=0" target="_blank"><font size=3><%=trim(RSreg("sTitle"))%> </font></a></td>
	<td nowrap><span class="eTableContent"><a href="<%=trim(RSreg("picPath"))%>" target="_blank"><img src="<%=trim(RSreg("picPath"))%>" alt="<%=trim(RSreg("picTitle"))%>" width="101" height="66" border="0"></a></span></td>
	<td nowrap><span class="eTableContent"><a href="cp_question.asp?iCUItem=<%=trim(RSreg("iCUItem"))%> "><%=trim(RSreg("parentIcuitem"))%></a></td>

	<td align="center" nowrap><span class="eTableContent">
	  <INPUT name="button" type="button" value="通過" onClick="PassValidate('<%=RSreg("picId")%>','<%=RSreg("parentIcuitem")%>','<%=totRec%>')">
      <INPUT name="button" type="button" value="不通過" onClick="NoPassValidate('<%=RSreg("picId")%>','<%=RSreg("parentIcuitem")%>','<%=totRec%>')">
	</span></td>
	<td align="center" nowrap>
	<span class="style1">
	<%if RSreg("picStatus")="W" then%>
	待審核
	<%end if%>
	</span>
	<%if RSreg("picStatus")="Y" then%>
	通過
	<%end if%>
	<%if RSreg("picStatus")="N" then%>
	不通過
	<%end if%>
	</td>
</tr>
                      
<% 
 RSreg.movenext 
  wend
%>
</TABLE>
           <div align="center">
             <input type="button" name="button" value="整批通過" onClick="AllPassValidate()"  class="cbutton" />
             <input type="button" name="button" value="整批不通過" onClick="AllNoPassValidate()" class="cbutton" />   
           </div>
	
       <!-- 程式結束 ---------------------------------->  
  <script type="text/javascript">
    function showSTitle(title)
    {
        alert((title));
    }
  </script>       
</form>  
</body>
</html>     
<%  end if%>
<script language="javascript">
	function PassValidate(id,parentIcuitem,totalRecord) 
	{
	    document.getElementById("picId").value = id;
		document.getElementById("parentIcuitem").value = parentIcuitem;
		document.getElementById("totalRecord").value = totalRecord;
		document.getElementById("submitTask").value = "Pass";
		document.getElementById("reg").submit();
	}
	function NoPassValidate(id,parentIcuitem,totalRecord) {
	    document.getElementById("picId").value = id;
		document.getElementById("parentIcuitem").value = parentIcuitem;
		document.getElementById("totalRecord").value = totalRecord;
		document.getElementById("submitTask").value = "NoPass";
		document.getElementById("reg").submit();
	}
	function AllPassValidate() 
	{
	    var obj=document.getElementsByName("checkbox");
        var len = obj.length;
        var checked = false;
        var CheckList="";
        for (i = 0; i < len; i++)
        {		
            if (obj[i].checked == true)
            {
			  CheckList+=obj[i].id + ";" ;
            }
        } 
	    document.getElementById("selectedItems").value = CheckList;	
		document.getElementById("submitTask").value = "AllPass";
		document.getElementById("reg").submit();
	}
	function AllNoPassValidate() 
	{
	    var obj=document.getElementsByName("checkbox");
        var len = obj.length;
        var checked = false;
        var CheckList="";
        for (i = 0; i < len; i++)
        {		
            if (obj[i].checked == true)
            {
			  CheckList+=obj[i].id + ";" ;
            }
        } 
	    document.getElementById("selectedItems").value = CheckList;	
		document.getElementById("submitTask").value = "AllNoPass";
		document.getElementById("reg").submit();
	}
	
</script>     
<script Language=VBScript>
	
		sub formModSubmit()
			reg.submitTask.value = "DELETE"
			reg.action = "<%=HTProgPrefix%>List.asp?icuitem=<%=request.querystring("icuitem")%>"
			reg.Submit
		end sub
		
		dim gpKey
    sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>.asp?keep=Y&nowPage=" & newPage & "&pagesize=<%=PerPageSize%>"_
		   & "&validate=<%=trim(request("validate"))%>&picTitle=<%=Server.URLEncode(trim(request("picTitle")))%>&memberId=<%=Server.URLEncode(memberId)%>&PicidTerms=<%=request("PicidTerms")%>"    
	
	end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>.asp?keep=Y&nowPage=<%=nowPage%>" & "&pagesize=" & newPerPage  _
		   & "&validate=<%=trim(request("validate"))%>&picTitle=<%=Server.URLEncode(trim(request("picTitle")))%>&memberId=<%=Server.URLEncode(memberId)%>&PicidTerms=<%=request("PicidTerms")%>"                 
     end sub 

     sub shortLongList(sorl)
           document.location.href="<%=HTProgPrefix%>.asp?keep=Y&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>" _
			& "&shortLongList=" & sorl
     end sub   

	</script>                                       
                       

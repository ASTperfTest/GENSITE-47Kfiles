
<% 
HTProgCode="GC1AP9"
HTProgPrefix="newkpi" 
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgPrefix = "lp_knowledgepic"
%>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/server.inc" -->
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

<%
picId=request("picId")
parentIcuitem=request("parentIcuitem")

		'---更新圖片狀態---
if request("submitTask") = "Pass" then
	sql = "UPDATE KnowledgePicInfo SET picStatus = 'Y' WHERE parentIcuitem = "& "'"  & parentIcuitem & "'" & " and picId= "& picId 
	
	conn.execute(sql)

elseif request("submitTask") = "NoPass" then

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


 
cSql = "SELECT     COUNT(*) AS allcount FROM  KnowledgePicInfo INNER JOIN CuDTGeneric ON KnowledgePicInfo.parentIcuitem = CuDTGeneric.iCUItem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account"

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
  '---狀態選擇時--
  validate=request("validate")

select case validate
	case "W"
	fSql= fsql & " * FROM KnowledgePicInfo INNER JOIN CuDTGeneric ON KnowledgePicInfo.parentIcuitem = CuDTGeneric.iCUItem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account ORDER BY CuDTGeneric.xPostDate DESC"
    fSql= "SELECT TOP "& nowPage * PerPageSize  & fSql & "where picStatus = 'W'"
	case "Y"
	fSql= fsql & " * FROM KnowledgePicInfo INNER JOIN CuDTGeneric ON KnowledgePicInfo.parentIcuitem = CuDTGeneric.iCUItem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account ORDER BY CuDTGeneric.xPostDate DESC"
    fSql= "SELECT TOP "& nowPage * PerPageSize  & fSql & "where picStatus = 'Y'"
	case "N"
	fSql= fsql & " * FROM KnowledgePicInfo INNER JOIN CuDTGeneric ON KnowledgePicInfo.parentIcuitem = CuDTGeneric.iCUItem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account ORDER BY CuDTGeneric.xPostDate DESC"
    fSql= "SELECT TOP "& nowPage * PerPageSize  & fSql & "where picStatus = 'N'"
    case else
    fSql= fsql & " * FROM KnowledgePicInfo INNER JOIN CuDTGeneric ON KnowledgePicInfo.parentIcuitem = CuDTGeneric.iCUItem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account ORDER BY CuDTGeneric.xPostDate DESC"
    fSql= "SELECT TOP "& nowPage * PerPageSize  & fSql 
end select 

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


<Form id="Form2" name=reg method="POST" action="lp_knowledgepiclist.asp">
 <INPUT TYPE=hidden name=submitTask value="">
 <INPUT TYPE=hidden name=picid value="">
  <INPUT TYPE=hidden name=parentIcuitem value="">
  <INPUT TYPE=hidden name=selectedItems value="">
<div class="browseby">
條件篩選依： 審核狀態
  <select name="validate" id="validate" >
  		
    <option value="" <%if validate="" then%>selected<%END IF%>>請選擇</option>
    <option value="all" <%if validate="all" then%>selected<%END IF%>>全部</option>
    <option value="W" <%if validate="W" then%>selected<%END IF%>>待審核</option>
    <option value="Y" <%if validate="Y" then%>selected<%END IF%>>審核通過</option>
    <option value="N" <%if validate="N" then%>selected<%END IF%>>審核不通過</option>
  </select>
<!--｜ 會員身分
<select name="select2" id="select2">
  <option selected>請選擇</option>
  <option>全部</option>
  <option>一般會員</option>
  <option>學者會員</option>
  <option>專家</option>
</select-->
</div>
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
			<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>">上一頁</a> ｜
    <% end if %>跳至第
    <select id=GoPage size="1" style="color:#FF0000" class="select">
		<% For iPage=1 to totPage %> 
			<option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
    <% Next %>   
    </select>      
		頁	
	<% if cint(nowPage)<>totPage then %> 
     ｜<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>">下一頁
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
					<th class=eTableLable>圖片標題</th>
					<th nowrap scope="col">圖片</th>
					<th nowrap scope="col">圖片上傳資料ID</th>
					<th nowrap scope="col">審核</th>
					<th nowrap scope="col">狀態</th>
					<!--<th class="First" scope="col">預覽</th>   -->
	  </tr>     

<% while not RsREG.eof %>	  
<tr>
  <td align="center"><span class="eTableContent">
    <input type="checkbox" name="checkbox"  id="<%=RSreg("picId")%>" value="checkbox">
  </span></td>                  
	   
	<td align="center"><%=trim(RSreg("account"))%></td>
	<td align="center"><%=trim(RSreg("realname"))%></td>
	<td align="center"><%=trim(RSreg("nickname"))%></td>
	<TD class=eTableContent><font size=3><%=trim(RSreg("picTitle"))%> </font></td>
	<td nowrap><span class="eTableContent"><a href="<%=trim(RSreg("picPath"))%>" target="_blank"><img src="<%=trim(RSreg("picPath"))%>" alt="圖片" width="101" height="66" border="0"></a></span></td>
	<td nowrap><a href="cp_question.asp?iCUItem=<%=trim(RSreg("iCUItem"))%> "><%=trim(RSreg("parentIcuitem"))%></a></td>
	<td align="center" nowrap><span class="eTableContent">
	  <INPUT name="button" type="button" value="通過" onClick="PassValidate('<%=RSreg("picId")%>','<%=RSreg("parentIcuitem")%>')">
      <INPUT name="button" type="button" value="不通過" onClick="NoPassValidate('<%=RSreg("picId")%>','<%=RSreg("parentIcuitem")%>')">
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
</form>  
</body>
</html>     

<script language="javascript">
	function PassValidate(id,parentIcuitem) 
	{
	    document.getElementById("picId").value = id;
		document.getElementById("parentIcuitem").value = parentIcuitem;
		document.getElementById("submitTask").value = "Pass";
		document.getElementById("reg").submit();
	}
	function NoPassValidate(id,parentIcuitem) {
	    document.getElementById("picId").value = id;
		document.getElementById("parentIcuitem").value = parentIcuitem;
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
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=" & newPage & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>" & "&pagesize=" & newPerPage                    
     end sub 

     sub shortLongList(sorl)
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>" _
			& "&shortLongList=" & sorl
     end sub   

	 
	 sub validate_OnChange
		document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>" _
			& "&validate=" & reg.validate.value 
	 end sub
	</script>                                       
                       

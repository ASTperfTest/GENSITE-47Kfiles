<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="單元資料維護"
HTProgFunc="新增"
HTUploadPath=session("Public")+"data/"
HTProgCode="GC1AP1"
HTProgPrefix="DsdXML" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
'response.write "xx="+request("CtNodeIDStr")
'response.end
    Set RSreg = Server.CreateObject("ADODB.RecordSet")
    nowPage = Request.QueryString("nowPage")  '現在頁數  

    if nowpage="" then
	fSql = "SELECT CG.iCUItem,CG.sTitle,CG.xKeyword,CG.created_Date,CG.xPostDate, htx.CtUnitName, " & _
		"xref1.mValue AS xref1CtUnitKind " & _
		"From CatTreeNode AS CTN " & _
		"LEFT JOIN CtUnit AS htx ON CTN.CtUnitID=htx.CtUnitID " & _
		"LEFT JOIN CuDTGeneric AS CG ON CG.iCTUnit=htx.CtUnitID " & _
		"LEFT JOIN CodeMainLong AS xref1 ON xref1.mCode = htx.CtUnitKind AND xref1.codeMetaID='refCTUKind' " & _
		"LEFT JOIN CodeMain AS xref3 ON xref3.mCode = htx.fCtUnitOnly AND xref3.codeMetaID='boolYN' " & _
		"WHERE CG.fCTUPublic='Y' " 
	if request("CtNodeIDStr")<>"" then
		fSql=fSql&" AND CTN.CtNodeID in ("+request("CtNodeIDStr")+")"	
	end if
	if request("htx_CtUnitKind")<>"" then
		fSql=fSql&" AND htx.CtUnitKind = '"+request("htx_CtUnitKind")+"'"
	end if
	if request("htx_CtUnitName")<>"" then
'		fSql=fSql&" AND htx.CtUnitName Like '%"+request("htx_CtUnitName")+"%'"
	end if	
	if request("htx_Stitle")<>"" then
		fSql=fSql&" AND CG.sTitle Like '%"+request("htx_Stitle")+"%'"
	end if	
	if request("htx_xKeyword")<>"" then
		fSql=fSql&" AND CG.xKeyword Like '%"+request("htx_xKeyword")+"%'"
	end if	
	if request("htx_IDateS") <> "" then
			rangeS = request("htx_IDateS")
			rangeE = request("htx_IDateE")
			if rangeE = "" then	rangeE=rangeS
		fSql = fSql & " AND CG.created_Date BETWEEN '"+rangeS+"' and '"+rangeE+"'"
	end if
	fSql = fSql & " ORDER BY CG.xPostDate DESC,CG.iCTUnit "
	session("refIDSQL") = fSql	
    else
    	fSql = session("refIDSQL")
    end if	
'    response.write fSql
'    response.end

'----------HyWeb GIP DB CONNECTION PATCH----------
'    RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------

    
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=15
      end if 
      
      RSreg.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
   end if   

'response.write cStr(nowPage)+"[]"+cStr(totRec)+"[]"+cStr(PerPageSize)+"[]"+cStr(totPage)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>選取引用資料清單</title>
<link href="css/popup.css" rel="stylesheet" type="text/css">
<link href="css/editor.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="PopFormName">單元資料維護 - 選取引用資料查詢清單</div>
<form action="DsdXMLAdd_refID2.asp" method="POST" id="PopForm">
引用方式: <input name=showType type="radio" value="4" checked>複製<input name=showType type="radio" value="5">參照&nbsp;&nbsp;&nbsp;

<input type="button" class="InputButton" value="重設查詢條件" onClick="Javascript:window.navigate('DsdXMLAdd_refID.asp')">
<input type="button" class="InputButton" value="回新增主畫面" onClick="Javascript:window.navigate('DsdXMLAdd.asp')">
</form>
<form name="form1" method="POST" action="">
<table border="0" width="500" cellspacing="0" cellpadding="0">
   <tr>
    <td align=center> 
     <font size="2" color="rgb(63,142,186)"> 第
     <font size="2" color="#FF0000"><%=nowPage%>/<%=totPage%></font>頁|                      
        共<font size="2" color="#FF0000"><%=totRec%></font>
       <font size="2" color="rgb(63,142,186)">筆| 跳至第       
         <select id=GoPage size="1" style="color:#FF0000">
             <%For iPage=1 to totPage%> 
                   <option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
             <%Next%>   
         </select>      
         頁</font>           
       <% if cint(nowPage) <>1 then %>             
            |<a href="DsdXMLAdd_refID2.asp?nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>">上一頁</a> 
       <%end if%>      
       
        <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %> 
            |<a href="DsdXMLAdd_refID2.asp?nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>">下一頁</a> 
        <%end if%>     
        | 每頁筆數:
       <select id=PerPage size="1" style="color:#FF0000">        
             <option value="99999"<%if PerPageSize=15 then%> selected<%end if%>>15</option>                                            
             <option value="100"<%if PerPageSize=30 then%> selected<%end if%>>30</option>                       
             <option value="200"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
             <option value="300"<%if PerPageSize=300 then%> selected<%end if%>>300</option>
        </select>     
     </font></td>
    <td align=right> <input type="button" class="InputButton" value="確定" onClick="closeForm()">
    </td>       
    </tr>
</table> 
<table width="100%" cellspacing="0" id="PopList">
  <tr>
	<th scope="col">&nbsp;</td>
	<th scope="col">標題</td>
	<th scope="col">主題單元名稱</td>
	<th scope="col">單元類型</td>
	<th scope="col">張貼日期</td>
	<th scope="col">關鍵字詞</td>
	<th scope="col">建立日期</td>
  </tr>
<%
    if not RSreg.EOF then              
    	for i=1 to PerPageSize                  	
%>
  <tr>
    <td class="Center">     <input type="radio" value="<%=RSreg("iCuItem")%>" name="iCuItem"></td>
    <td><%=RSreg("sTitle")%></td>
    <td><%=RSreg("CtUnitName")%></td>
    <td><%=RSreg("xref1CtUnitKind")%></td>
    <td><%=RSreg("xPostDate")%></td>
    <td><%=RSreg("xKeyword")%>&nbsp;</td>
    <td><%=RSreg("created_Date")%>&nbsp;</td>
  </tr>
<%
             RSreg.moveNext
             if RSreg.eof then exit for 
    	next 
    end if    %>
</table>
</form>
</body>
</html>
<script language=vbs>
sub GoPage_OnChange      
	newPage=form1.GoPage.value     
        document.location.href="DsdXMLAdd_refID2.asp?nowPage=" & newPage & "&pagesize=<%=PerPageSize%>"    
end sub      
     
sub PerPage_OnChange                
        newPerPage=form1.PerPage.value
        document.location.href="DsdXMLAdd_refID2.asp?nowPage=<%=nowPage%>&pagesize=" & newPerPage                    
end sub       

sub closeForm()
        radiochk=false
  	for i=0 to form1.iCuItem.length-1                           
   	    if form1.iCuItem(i).checked then 
      		radiochk=true                          
      		exit for                           
    	    end if                           
  	next  
  	if radiochk then
  	    if PopForm.showType(0).checked then
		window.navigate "DsdXMLAdd.asp?showType=4&iCuItem="& form1.iCuItem(i).value 
	    else
		window.navigate "DsdXMLAdd_showType5.asp?showType=5&iCuItem="& form1.iCuItem(i).value 	    
	    end if
	else
		alert "請選取資料!"
		exit sub
	end if
end sub

</script>


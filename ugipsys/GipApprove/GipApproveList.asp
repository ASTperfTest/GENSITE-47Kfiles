<%@ CodePage = 65001 %>
<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="資料審稿"
if request("ctNodeId")="" then
	HTProgFunc="待審清單"
else
	HTProgFunc="資料清單"
end if
HTProgCode="GC1AP2"
HTProgPrefix="GipApprove" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
if request("ctRootId")<>"" then session("ctRootId") = request("ctRootId")
if request("submitTask")<>"" then
    SQLUpdate=""
    For Each x In Request.Form
	if left(x,5)="ckbox" and request(x)<>"" then
	    xn=mid(x,6)
	    SQLUpdate=SQLUpdate+"Update CuDtGeneric Set fctupublic=N'"&request("fctupublic")&"', " & _
	    	"ieditor=N'"&session("userId")&"', deditDate=getdate() " & _
	    	"where icuitem=" & request("xphKeyicuitem"&xn) & ";"
        end if
    Next 
    if SQLUpdate<>"" then conn.execute(SQLUpdate)
%>
	<script language=VBS>
		alert "審核完成！"
		window.navigate "<%=HTProgPrefix%>List.asp?ctRootId=<%=request("ctRootId")%>&ctNodeId=<%=request("ctNodeId")%>&nowpage=1"
	    </script>
<%	   
	response.end
else
    nowPage=Request.QueryString("nowPage")  '現在頁數
    if nowpage="" then
    	fSql="Select C.icuitem,C.stitle,C.xkeyword,CT.ctUnitName,CM.mvalue xreftopDataCat,U.userName," _
    		& " D.deptName,C.deditDate,CTN.catName,C.fctupublic,CM2.mvalue xreffctupublic,CTR.ctRootName,CTR.pvXdmp " & _
			" from CuDtGeneric C " & _
			" Left Join CatTreeNode CTN ON C.ictunit=CTN.ctUnitId " & _
			" Left Join CatTreeRoot CTR ON CTN.ctRootId=CTR.ctRootId " & _
			" Left Join CtUserSet2 CUS2 ON CTN.ctNodeId=CUS2.ctNodeId AND CUS2.userId=N'"&session("userId")&"' " & _
			" Left Join CtUnit CT On C.ictunit=CT.ctUnitId " & _
			" Left Join CodeMain CM ON C.topCat=CM.mcode AND CM.codeMetaId=N'topDataCat' " & _
			" Left Join CodeMain CM2 ON C.fctupublic=CM2.mcode AND CM2.codeMetaId=N'isPublic3' " & _
			" Left Join Dept D On C.idept=D.deptId " & _
			" Left Join InfoUser U On C.iEditor=U.userID " & _
			" where CUS2.rights=1 "
		if request("ctNodeId")<>"" then
			fSql=fSql&" AND CTN.ctNodeId =N'"&request("ctNodeId")&"'"
		else
			if request("htx_fctupublic")="" then
				fSql=fSql&" AND (C.fctupublic<>N'Y' AND C.fctupublic<>N'N')"
			else
				fSql=fSql&" AND C.fctupublic=" & pkStr(request("htx_fctupublic"),"")
			end if
		end if
		if request("htx_ctRootID")<>"" then
			fSql=fSql&" AND CTN.ctRootId ="&request("htx_ctRootID")
		end if
		if request("htx_catName")<>"" then
			fSql=fSql&" AND CTN.catName Like N'%"+request("htx_catName")+"%'"
		end if
		if request("htx_stitle")<>"" then
			fSql=fSql&" AND C.stitle Like N'%"+request("htx_stitle")+"%'"
		end if
		if request("htx_xBody")<>"" then
			fSql=fSql&" AND C.xBody Like N'%"+request("htx_xBody")+"%'"
		end if
		if request("htx_IDateS") <> "" then
				rangeS = request("htx_IDateS")
				rangeE = request("htx_IDateE")
				if rangeE = "" then	rangeE=rangeS
			fSql = fSql & " AND C.deditDate BETWEEN N'"+rangeS+"' and N'"+rangeE+"'"
		end if
		if (HTProgRight AND 64) = 0 then
			fSql = fSql & " AND C.idept=" & pkStr(session("deptId"),"")
		end if
		fSql = fSql & " Order by CTN.ctRootId,C.deditDate DESC"
		session("GipApproveList")=fSql
'		response.write fsql & "<HR>"
'		response.end
    else
    	fSql=session("GipApproveList")
    end if
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=15  
      end if 

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
'	RSreg.CacheSize = PerPageSize

	

'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,3,1
Set RSreg = Conn.execute(fSql)
	RSreg.CacheSize = PerPageSize
'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      
      RSreg.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
   end if    
end if   
end if
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
<script Language=VBScript>
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=" & newPage & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=<%=nowPage%>&pagesize=" & newPerPage                    
     end sub 

      Dim chkCount
      chkCount=0            '記錄checkbox 被勾數

    sub document_onClick           'checkbox 被勾計數
         set sObj=window.event.srcElement
         if sObj.tagName="INPUT" then 
            if sObj.type="checkbox"  then 
                if sObj.checked then 
                   chkCount=chkCount+1
                else
                   chkCount=chkCount-1                
                end if                                          
            end if
         end if
         '
         if chkCount=0 then 
            document.all("RunJob").style.visibility="hidden"
         else
            document.all("RunJob").style.visibility="visible"
         end if
    end sub        
    sub Chkall
       chkCount=0     
       if document.all("ckall").value="全選" then           '全勾
          for i=0 to reg.elements.length-1
               set e=document.reg.elements(i)
               if left(e.name,5)="ckbox" then 
                   e.checked=true
                   chkCount=chkCount+1
              end if     
          next                 
          document.all("RunJob").style.visibility="visible"
          document.all("ckall").value="全不選"
      elseif document.all("ckall").value="全不選" then        '全不勾
          for i=0 to reg.elements.length-1
               set e=document.reg.elements(i)
               if left(e.name,5)="ckbox" then 
                   e.checked=false
               end if     
          next                 
          document.all("RunJob").style.visibility="hidden"
          document.all("ckall").value="全選"          
       end if
   end sub   
   sub button1_onClick
   	if document.all("sfx_fctupublic").value="" then
   		alert "請點選是否公開狀態欄位!"
   		document.all("sfx_fctupublic").focus
   		exit sub
   	end if
   	xPos=instr(document.all("sfx_fctupublic").value,"--")
   	fctupublicValue=Left(document.all("sfx_fctupublic").value,xPos-1)
   	fctupublicDisplay=mid(document.all("sfx_fctupublic").value,xPos+2)
        chky=msgbox("注意！"& vbcrlf & vbcrlf &"　您確定要將資料狀態變更為 [" &fctupublicDisplay& "] 嗎？"& vbcrlf , 48+1, "請注意！！")
        if chky=vbok then
              document.all("fctupublic").value=fctupublicValue
	      reg.submitTask.value = "UPDATE"
	      reg.Submit
       	end If   	
   end sub

</script>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
 <Form name=reg method="POST" action=<%=HTprogPrefix%>List.asp>
 <INPUT TYPE=hidden name=submitTask value="">
 <INPUT TYPE=hidden name=ctRootId value="<%=request("ctRootId")%>"> 
 <INPUT TYPE=hidden name=ctNodeId value="<%=request("ctNodeId")%>"> 
 <INPUT TYPE=hidden name=fctupublic value=""> 
  <tr>
	    <td class="FormName"><%=HTProgCap%>&nbsp;
	    <font size=2>【<%=HTProgFunc%>】</td>
    	<td class="FormLink" valign="top" align=right>
		<%if (HTProgRight and 2)=2 then%>
			<A href="GipApproveQuery.asp" title="指定查詢條件">待審清單查詢</A>
		<%end if%>

    	</td>    	    
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top" align=right colspan="2">
    	 公開狀態: <Select name="sfx_fctupublic" size=1>
         	<option value="">請選擇</option>
 			<%SQL="Select mcode,mvalue from CodeMain where codeMetaId=N'isPublic3' Order by msortValue DESC"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>--<%=RSS(1)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
        	
         </select>
	   <SPAN id=RunJob style="visibility:hidden">
	   	   <input type=button class=cbutton value="確定" id=button1 name=button1>
	   </SPAN>
    </td>    
    </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
  <p align="center">  
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
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>">上一頁</a> 
       <%end if%>      
       
        <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %> 
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>">下一頁</a> 
        <%end if%>     
        | 每頁筆數:
       <select id=PerPage size="1" style="color:#FF0000">            
             <option value="15"<%if PerPageSize=15 then%> selected<%end if%>>15</option>                       
             <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
             <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
             <option value="300"<%if PerPageSize=300 then%> selected<%end if%>>300</option>
        </select>     
     </font>     
    <CENTER>
     <TABLE width=100% cellspacing="1" cellpadding="0" class=bg>                   
     <tr align=left>    
	<td align=center class=eTableLable width=7%>
	<input type=button value ="全選" class="cbutton"  name=ckall onClick="ChkAll"></td>
     	<td class=eTableLable>狀態</td>
     	<td class=eTableLable>預覽</td>
     	<td class=eTableLable>目錄樹</td>
     	<td class=eTableLable>單元名稱</td>
     	<td class=eTableLable>標  題</td>
     	<td class=eTableLable>帳號</td> 
     	<td class=eTableLable>單位</td> 
     	<td class=eTableLable>編修日期</td> 
     	<td class=eTableLable>編修</td>      	
     </tr> 	     	     	 
<%
If not RSreg.eof then   

    for i=1 to PerPageSize
    	xUrl = session("myWWWSiteURL") & "/content.asp?mp=" & RSreg("pvXdmp") & "&CuItem="  & RSreg("icuitem")
    	pKey = "icuitem=" & RSreg("icuitem")
    	if RSreg("fctupublic")="P" then
    		fctupublicStr="<font color=red>"&RSreg("xreffctupublic")&"</font>"
    	else
    		fctupublicStr=RSreg("xreffctupublic")
    	end if
%>                  
<tr>                  
<TD align=center><INPUT TYPE=checkbox name="ckbox<%=i%>" value="Y">
<INPUT TYPE=hidden name="xphKeyicuitem<%=i%>" value="<%=RSreg("icuitem")%>"></td>   
	<TD class=eTableContent><font size=2>
	<%=fctupublicStr%></TD>	
	<TD class=eTableContent><font size=2>
	<A href="<%=xUrl%>" target="_wMof">View</A>&nbsp;</TD>
	<TD class=eTableContent><font size=2>
	<%=RSreg("ctRootName")%></TD>	
	<TD class=eTableContent><font size=2>
	<%=RSreg("catName")%></TD>	
	<TD class=eTableContent><font size=2><%=RSreg("stitle")%>
	</TD>
	<TD class=eTableContent><font size=2>
	<%=RSreg("UserName")%></TD>			
	<TD class=eTableContent><font size=2>
	<%=RSreg("deptName")%></TD>			
	<TD class=eTableContent><font size=2>
	<%=xStdTime(RSreg("deditDate"))%></TD>
	<td align=center class=eTableContent>
	<input type=button value ="編修" class="cbutton" onClick="DsdXMLEdit '<%=RSreg("icuitem")%>'"></td>				
    </tr>
    <%
         RSreg.moveNext
         if RSreg.eof then exit for 
    next 
   %>
<%else%>
    <tr>                  
    <TD align=center class=eTableContent colspan=7><font size=2>目前無待審資料!</font>
    </td></tr>
<%end if%>

    </TABLE>    
    </CENTER>
     </form>      
    </td>
  </tr>    
</table>   
</body>
</html>  
<script language=VBS>
sub DsdXMLEdit(icuitem)
'document.location.reload
	window.open "../GipEdit/DsdXMLEdit.asp?phase=edit&S=Approve&icuitem="&icuitem,"","scrollbars=yes,width=800,height=600"
end sub
sub GipApproveListRefresh()
	document.location.reload
end sub
</script>



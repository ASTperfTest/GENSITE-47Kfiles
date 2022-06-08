<%@ codepage=65001 %>
<% Response.Expires = 0
HTProgCap="GIP資料匯出"
HTProgFunc="清單"
HTProgPrefix="GIPDataXMLGen" %>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

function xStdTime2(dt)    
   if Len(dt)=0 or isnull(dt) then
     	xStdTime2=""
   else
        xyear = cstr(year(dt))     '補零
        xmonth = right("00"+ cstr(month(dt)),2)
        xday = right("00"+ cstr(day(dt)),2)
        xhour = right("00" + cstr(hour(dt)),2)
        xminute = right("00" + cstr(minute(dt)),2)
        xsecond = right("00" + cstr(second(dt)),2)
        xStdTime2 = xyear & "/" & xmonth & "/" & xday & " " & xhour & ":" & xminute & ":" & xsecond
   end if
end function

HTUploadPath=session("Public")+"GIPDataXML"	
sourcePath = server.mappath(session("Public")+"data") & "\"
sourcePath2 = server.mappath(session("Public")+"Attachment") & "\"
'response.write sourcePath & "<br>"
'response.write sourcePath2 & "<br>"
'response.write server.MapPath(HTUploadPath)

if request("submitTask")<>"" then
	ct=chr(9)
	'----處理DTD或DSD XML
	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true		
    '----找出對應的DSD,若不存在則用DTD
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filePath = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & request("CtUnitID") & ".xml")
	if fso.FileExists(filePath) then
		LoadXML = filePath
	else
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & request("iBaseDSD") & ".xml")
	end if     	
	xv = htPageDom.load(LoadXML)	
  	if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reason: " &  htPageDom.parseError.reason)
    		Response.End()
  	end if
	set refModel = htPageDom.selectSingleNode("//dsTable")
	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")     	
'	response.write "<XMP>"+allModel.xml+"</XMP>"	
'	response.end
	BaseDSDTableName = nullText(refModel.selectSingleNode("tableName"))
'	response.write BaseDSDTableName
'	response.end
	XMLPath = HTUploadPath+"/OUTXML/"+request("CtUnitID")+".xml"
	Set xfout = fso.CreateTextFile(server.mapPath(XMLPath),true,true)
	xfout.writeline "<?xml version=""1.0""  encoding=""utf-16"" ?>"
	xfout.writeline "<GIPDataXML>"	
	xfout.writeline ct & "<GIPData>"	
    SQLXML=""
'    on error resume next
	outputCount = 0
    For Each x In Request.Form
		if left(x,5)="ckbox" and request(x)<>"" then
			xfout.writeline ct & ct & "<fieldList>"	
		    xn=mid(x,6)
		    SQLXML="Select * from CuDTGeneric htx Left Join " & BaseDSDTableName & " ghtx ON htx.iCuItem=ghtx.giCuItem " & _
		    	" where htx.iCUItem=" & request("xphKeyiCUItem"&xn) & ";"
		    Set RSXML=conn.execute(SQLXML)
			'----Master處理
			for each param in allModel.selectNodes("dsTable[tableName='CuDtGeneric']/fieldList/field[fieldName!='icuitem' and fieldName!='ibaseDsd' and fieldName!='ictunit']") 
				xfout.writeline ct & ct & ct & "<field>"	
				xfout.writeline ct & ct & ct & ct & "<fieldName>"+nullText(param.selectSingleNode("fieldName"))+"</fieldName>"	
				xStr = ""
				if param.selectSingleNode("fieldName").text="createdDate" or param.selectSingleNode("fieldName").text="deditDate" then				
					if not isNull(RSXML(nullText(param.selectSingleNode("fieldName")))) then 				
						xStr = xStdTime2(RSXML(nullText(param.selectSingleNode("fieldName"))))
					else				
						xStr = ""
					end if
				else
					if not isNull(RSXML(nullText(param.selectSingleNode("fieldName")))) then 				
						xStr = RSXML(nullText(param.selectSingleNode("fieldName")))
					else				
						xStr = ""
					end if				
				end if
				xfout.writeline ct & ct & ct & ct & "<fieldValue><![CDATA["+cStr(xStr)+"]]></fieldValue>"	
				xfout.writeline ct & ct & ct & "</field>"	
	    		'----主圖與檔案下載式檔案複製
	    		if not isNull(RSXML("xImgFile")) then
					if fso.FileExists(sourcePath+RSXML("xImgFile")) then
						fso.CopyFile sourcePath+RSXML("xImgFile"),server.mapPath(HTUploadPath+"/Data")+"\"+RSXML("xImgFile")
					end if
			
	end if
	    		if not isNull(RSXML("fileDownLoad")) then
					if fso.FileExists(sourcePath+RSXML("fileDownLoad")) then
						fso.CopyFile sourcePath+RSXML("fileDownLoad"),server.mapPath(HTUploadPath+"/Data")+"\"+RSXML("fileDownLoad")
					end if
			
	end if		
			next 			
			'----Detail表處理
			for each param in refModel.selectNodes("fieldList/field")
				xfout.writeline ct & ct & ct & "<field>"	
				xfout.writeline ct & ct & ct & ct & "<fieldName>"+nullText(param.selectSingleNode("fieldName"))+"</fieldName>"	
				xStr = ""
				if not isNull(RSXML(nullText(param.selectSingleNode("fieldName")))) then 
					xStr = RSXML(nullText(param.selectSingleNode("fieldName")))
				else				
					xStr = ""
				end if
				xfout.writeline ct & ct & ct & ct & "<fieldValue><![CDATA["+cStr(xStr)+"]]></fieldValue>"	
				xfout.writeline ct & ct & ct & "</field>"	
			next  
			'----準備附件XML
			SQLA = "Select CDA.*,I.OldFileName from CuDTAttach CDA Left Join ImageFile I ON CDA.NFileName=I.NewFileName " & _
				"where CDA.bList='Y' and CDA.xiCuItem=" & pkStr(request("xphKeyiCUItem"&xn),"")
			Set RSA =  conn.execute(SQLA)
			if not RSA.eof then
				xfout.writeline ct & ct & ct & "<attachList>"	
			    while not RSA.eof
					xfout.writeline ct & ct & ct & ct & "<Attach>"	
					xfout.writeline ct & ct & ct & ct & ct & "<aTitle><![CDATA["+RSA("aTitle")+"]]></aTitle>"	
					xfout.writeline ct & ct & ct & ct & ct & "<aDesc><![CDATA["+RSA("aDesc")+"]]></aDesc>"	
					xfout.writeline ct & ct & ct & ct & ct & "<NFileName>"+RSA("NFileName")+"</NFileName>"	
					xfout.writeline ct & ct & ct & ct & ct & "<OldFileName>"+RSA("OldFileName")+"</OldFileName>"	
					xfout.writeline ct & ct & ct & ct & ct & "<bList>"+RSA("bList")+"</bList>"	
					xfout.writeline ct & ct & ct & ct & ct & "<aEditor>"+RSA("aEditor")+"</aEditor>"	
					xfout.writeline ct & ct & ct & ct & ct & "<aEditDate>"+cStr(RSA("aEditDate"))+"</aEditDate>"	
					xfout.writeline ct & ct & ct & ct & ct & "<listSeq>"+RSA("listSeq")+"</listSeq>"							    	
					xfout.writeline ct & ct & ct & ct & "</Attach>"	
					'----複製附件檔案
		    		if not isNull(RSA("NFileName")) then
						if fso.FileExists(sourcePath2+RSA("NFileName")) then
							fso.CopyFile sourcePath2+RSA("NFileName"),server.mapPath(HTUploadPath+"/Attachment")+"\"+RSA("NFileName")
						end if
				
	end if							
			    	RSA.movenext
			    wend
				xfout.writeline ct & ct & ct & "</attachList>"	
			end if
			xfout.writeline ct & ct & "</fieldList>"	
			outputCount = outputCount + 1
        end if
    Next 
	xfout.writeline ct & "</GIPData>"	
	xfout.writeline "</GIPDataXML>"  
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
共寫入XML檔案 <%=outputCount%> 筆資料!<br/>
<A href="<%=XMLPath%>" target="_blank">檢查XML檔</A>&nbsp;&nbsp;&nbsp;&nbsp;
<A href="GIPDataXMLGen.asp">回匯出</A>
</body>
</html>  	
<%	
	response.end
else
	SQLC = "Select CtUnitName from CtUnit where CtUnitID=" & request("htx_CtUnitID")
	Set RSC = conn.execute(SQLC)
	if not RSC.eof then xCtUnitName = RSC(0)
    nowPage=Request.QueryString("nowPage")  '現在頁數
    if nowpage="" then
    	fSql="Select C.iCUItem,C.sTitle from CuDTGeneric C "
		fSql = fSql & " where iCtUnit='"&request("htx_CtUnitID")&"' Order by C.iCUItem"
		session("GipDataXMLGenList")=fSql
    else
    	fSql=session("GipDataXMLGenList")
    end if
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=9999 
      end if 
 	Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
'	RSreg.CacheSize = PerPageSize

'----------HyWeb GIP DB CONNECTION PATCH----------
'	RSreg.Open fSql,Conn,3,1
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
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
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
	      reg.submitTask.value = "UPDATE"
	      reg.Submit
   end sub

</script>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
 <Form name=reg method="POST" action=<%=HTprogPrefix%>List.asp>
 <INPUT TYPE=hidden name=submitTask value="">
 <INPUT TYPE=hidden name=iBaseDSD value="<%=request("htx_iBaseDSD")%>">
 <INPUT TYPE=hidden name=CtUnitID value="<%=request("htx_CtUnitID")%>">
 <INPUT TYPE=hidden name=CtUnitName value="<%=xCtUnitName%>">
  <tr>
	    <td class="FormName"><%=HTProgCap%>&nbsp;
	    <font size=2>【主題單元--<%=xCtUnitName%>】</td>
    	<td class="FormLink" valign="top" align=right>
    	</td>    	    
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top" align=right colspan="2">
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
             <option value="9999"<%if PerPageSize=9999 then%> selected<%end if%>>全部</option>
        </select>     
     </font>     
    <CENTER>
     <TABLE width=100% cellspacing="1" cellpadding="0" class=bg>                   
     <tr align=left>    
	<td align=center class=eTableLable width=7%>
	<input type=button value ="全選" class="cbutton"  name=ckall onClick="ChkAll"></td>
     	<td class=eTableLable>iCuItem</td>
     	<td class=eTableLable>標題</td>
     </tr> 	     	     	 
<%
If not RSreg.eof then   

    for i=1 to PerPageSize
%>                  
<tr>                  
<TD align=center><INPUT TYPE=checkbox name="ckbox<%=i%>" value="Y">
<INPUT TYPE=hidden name="xphKeyiCUItem<%=i%>" value="<%=RSreg("iCUItem")%>"></td>   
	<TD class=eTableContent><font size=2><%=RSreg("sTitle")%>
	</TD>
    </tr>
    <%
         RSreg.moveNext
         if RSreg.eof then exit for 
    next 
   %>
<%else%>
    <tr>                  
    <TD align=center class=eTableContent colspan=7><font size=2>目前無資料!</font>
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


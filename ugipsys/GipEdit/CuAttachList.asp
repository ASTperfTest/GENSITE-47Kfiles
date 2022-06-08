<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="資料附件"
HTProgFunc="條列"
HTProgCode="GC1AP1"
HTProgPrefix="CuAttach" %>
<% response.expires = 0 %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="../inc/setstyle.css">
<title></title>
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Dim mpKey, dpKey
Dim RSmaster, RSlist
	sqlCom = "SELECT * FROM CuDtGeneric WHERE icuitem=" & pkStr(request.queryString("icuitem"),"")
	Set RSmaster = Conn.execute(sqlcom)
	mpKey = ""
	mpKey = mpKey & "&icuitem=" & RSmaster("icuitem")
	if mpKey<>"" then  mpKey = mid(mpKey,2)
%>
<body>


<table border="0" width="95%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
			<A href="DsdXMLEdit.asp?<%=mpKey%>&phase=edit" title="Back">回資料編修</A>
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A> 
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
</table>
<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%"><TR><TD class="eTableLable" align="right"><font color="red">*</font>
標題：</TD>
<TD class="eTableContent"><input name="htx_sTitle" size="50">
<input type="hidden" name="htx_icuitem">
</TD>
</TR>
</TABLE>
</CENTER>
</form>
<%
function qqRS(fldName)
	xValue = RSmaster(fldname)
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		qqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
	end if
end function %>

<script language=vbs>
	reg.htx_sTitle.value= "<%=qqRS("sTitle")%>"
	reg.htx_icuitem.value= "<%=qqRS("icuitem")%>"
</script>
<%
	fSql = "SELECT dhtx.*, i.oldFileName " _
		& " FROM CuDtattach AS dhtx" _
		& " LEFT JOIN ImageFile AS i ON i.newFileName=dhtx.nfileName" _
		& " WHERE 1=1" _
		& " AND dhtx.xiCuItem=" & pkStr(RSmaster("icuitem"),"") _
		& " ORDER BY dhtx.listSeq"
		
	set RSlist = conn.execute(fSql)
%>

	<INPUT TYPE=button VALUE="新增附件" onClick="window.navigate('CuAttachAdd.asp?<%=mpKey%>&user=<%=session("userID")%>&path=..<% =session("Public") %>Attachment&siteid=<% =session("mySiteID") %>&phase=add')">

<CENTER>
 <TABLE width="95%" cellspacing="1" cellpadding="0" class="bg">
 <tr align="left">
	<td class=eTableLable>附件名稱</td>
	<td class=eTableLable>原檔名</td>
	<td class=eTableLable>顯示</td>
	<td class=eTableLable>次序</td>
	<td class=eTableLable>上傳人</td>
	<td class=eTableLable>上傳日</td>
 </tr>
<%
i=1
	while not RSlist.eof
		dpKey = ""
		dpKey = dpKey & "&ixCuAttach=" & RSlist("ixCuAttach")
		if dpKey<>"" then  dpKey = mid(dpKey,2)
		
%>
	<TD class=eTableContent><font size=2>

	<A href="CuAttachEdit.asp?<%=dpKey%>&user=<%=session("userID")%>&path=..<% =session("Public") %>Attachment&siteid=<% =session("mySiteID") %>&phase=edit">

<% =i %>.<%=RSlist("aTitle")%>
</A>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("oldFileName")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("bList")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("listSeq")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("aEditor")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("aEditDate")%>
</font></td>
    </tr>
    <%
         i=i+1  
         RSlist.moveNext
     wend
   %>
    </TABLE>
	
    </CENTER>
       <!-- 程式結束 ---------------------------------->  
     </form>  
	 
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center">
	</td></tr>
</table> 

<%
'Max  2011/05/24
'重新查詢，音樂及照片管理

'bob  2011/08/16
'只有styleG才出現
fSql_new =" SELECT "
fSql_new = fSql_new & vbcrlf & " 	 xiCuItem"
fSql_new = fSql_new & vbcrlf & "	,aTitle"
fSql_new = fSql_new & vbcrlf & "	,NFileName"
fSql_new = fSql_new & vbcrlf & "	,bList"
fSql_new = fSql_new & vbcrlf & "	,bgMusic "
fSql_new = fSql_new & vbcrlf & "FROM CuDtattach AS l "
fSql_new = fSql_new & vbcrlf & "JOIN CuDTx7 AS r ON l.xiCuItem=r.giCuItem "
fSql_new = fSql_new & vbcrlf & "WHERE bList = 'Y' "
fSql_new = fSql_new & vbcrlf & "AND right(NFileName, 4) in ('.jpg' ,'jpge' ,'.gif' ,'.bmp')"
fSql_new = fSql_new & vbcrlf & "AND xiCuItem=" & RSmaster("icuitem")
fSql_new = fSql_new & vbcrlf & "AND exists(select * from CatTreeNode where CTUnitId in ("
fSql_new = fSql_new & vbcrlf & "		select iCTUnit from CuDTGeneric where iCuItem=" & RSmaster("icuitem")
fSql_new = fSql_new & vbcrlf & "	) and xslData = 'styleG')"  '只有styleG才出現
fSql_new = fSql_new & vbcrlf & "ORDER BY listSeq "	

set RS_NewList = conn.execute(fSql_new)

sql_bgnusic = "Select * from BackgroundMusic"
set RS_bgMusic = conn.execute(sql_bgnusic)
%>
 <br />
 <% if not RS_NewList.eof then %>
<form method="POST" name="MusicMode" action="CuAttachModAct.asp">
<input type="hidden" name="iCuItemID" value="<%=request("icuitem")%>">
<% Function chk(val1,val2,strrtn) 
	   If UCase(val1)=UCase(val2) then
	   chk=strrtn
       Else
       chk=""
       End If
End Function 
%>
<fieldset style="font-size: 15px; text-align: left;">
            <legend>照片瀏覽設定(照片播放版型使用)</legend>
            <div style="margin-left: 20px"><br/>
                背景音樂設定：				
					<select name="sel" ><!--onchange="javascript:SelectMusic(this.options[this.options.selectedIndex].value);"-->> 
                    <option id="" value=""  <%=chk(RS_NewList("bgMusic"),"","selected")%>>無背景音樂</option>
                   <% while not RS_bgMusic.eof %>
                    <option id="<%=RS_bgMusic("bgMusicID") %>" value="<%=RS_bgMusic("bgMusicID") %>" <%=chk(RS_NewList("bgMusic"),RS_bgMusic("bgMusicID"),"selected")%>><%=RS_bgMusic("Title") %></option>
                    <%RS_bgMusic.moveNext
                    wend 
                    %>
                    <!--<option id="Music1" value="artist01.mp3"  <%=chk(RS_NewList("bgMusic"),"artist01.mp3","selected")%>>曲目1</option>
                    <option id="Music2" value="artist02.mp3"  <%=chk(RS_NewList("bgMusic"),"artist02.mp3","selected")%>>曲目2</option>
                    <option id="Music3" value="artist03.mp3"  <%=chk(RS_NewList("bgMusic"),"artist03.mp3","selected")%>>曲目3</option>
                    <option id="Music4" value="artist04.mp3"  <%=chk(RS_NewList("bgMusic"),"artist04.mp3","selected")%>>曲目4</option>-->
                    <option id="Music5" value="Random"  <%=chk(RS_NewList("bgMusic"),"Random","selected")%>>隨機</option>
					</select>  <input type="submit" value="更換">					  
					<br />
                <iframe scrolling="no" width="90%" border="0" frameborder="0" cellspacing="0" height="450" 
                src="<%=session("myWWWSiteURL") %>/subject/FileShowImage.aspx?CuItem=<%=RSmaster("icuitem") %>&T=<%=now() %>">			
				
            </div>
        </fieldset>
    </form>
	
	<% 	end if	%>
</body>
</html>                                 

<script Language=VBScript>
  dim gpKey
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=" & newPage & "&strSql=" & strSql & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=<%=nowPage%>" & "&strSql=<%=strSql%>&pagesize=" & newPerPage                    
     end sub 

     sub setpKey(xv)
     	gpKey = xv
     end sub

	sub butAction(k) 
	  if gpKey = "" then 
	    alert("請先選擇個案!") 
	  else
    	select case k 
    	end select 
	  end if 
	end sub 

</script>




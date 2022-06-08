<%@ CodePage = 65001 %>
<?xml version="1.0"  encoding="utf-8" ?>
<hpMain>
<MenuTitle /> 
 <CtUnitName>網站導覽</CtUnitName> 
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/xDataSet.inc" -->

<% 

Dim RSKey
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	

		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp" & request("mp") & ".xml"
		xv = htPageDom.load(LoadXML)
		'Response.write(server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp" & request("mp") & ".xml")
'		response.write xv & "<HR>"
  		if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    		Response.End()
  		end if

  		myTreeNode = 0
                upParent = 0

  	set refModel = htPageDom.selectSingleNode("//MpDataSet")
  	response.write "<myStyle>"&nullText(refModel.selectSingleNode("MpStyle"))&"</myStyle>"
	response.write "<mp>"&request("mp")&"</mp>"
	xRootID=nullText(refModel.selectSingleNode("MenuTree"))
	xSitemapID=nullText(refModel.selectSingleNode("MenuTree"))

%>	
	<!--#include file="gensite.inc" -->
  <!--#include file= "content.inc" -->
	<Sitemap myTreeNode="<%=myTreeNode%>">
<%

	SQLCom = "select * from CatTreeRoot Where CtRootID = "& xRootID
	set RS = conn.execute(SqlCom)


	FtypeName = RS("CtRootName")
	
	sqlstring = "select CatMemo_Disable from nodeinfo where ctrootid ="& xRootID
	set rsDis = conn.Execute(sqlstring)
	memo_disable = false
	if not rsDis.eof then
		if rsDis("CatMemo_Disable") <>"" or rsDis("CatMemo_Disable")<> NULL then
			memo_disable = rsDis("CatMemo_Disable")
		end if
	end if
   sql="select distinct datalevel from CatTreeNode where CtRootID = " & xRootID
   set rs0=conn.Execute(sql)
   do while not rs0.eof   
    if rs0(0)="1" then
        SqlCom = "SELECT a.CatName AS xCat,a.*,b.redirectURL, b.newWindow,b.iBaseDSD,b.CtUnitKind "_
                & " FROM CatTreeNode a,CtUnit b"_  
                & " WHERE  b.ctUnitID=*a.CtUnitID and a.inUse='Y' AND  a.DataLevel=" & rs0(0) & " AND  a.CtRootID = " & xRootID  _
                & " Order by  a.CatShowOrder"
   else
 	SqlCom = "SELECT b.CatName AS xCat, a.*, c.redirectURL, c.newWindow, c.iBaseDSD , c.CtUnitKind " _
		& " FROM CatTreeNode AS a JOIN CatTreeNode AS b ON b.CtNodeID=a.DataParent" _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.inUse='Y' AND a.DataLevel=" & rs0(0) & " AND a.CtRootID = " & xRootID  _
		& " Order by b.CatShowOrder, a.CatShowOrder"
    end if		
	'response.write SqlCom & "<HR>"

	set RS = conn.execute(SqlCom)
	xCat = ""
	while not RS.eof
	  if rs0(0)="1" then
		xCat = RS("xCat")
%>
           <Menucat iCuItem="<%=RS("CtNodeID")%>" dataparent="<%=RS("DataParent")%>" isUnit="<%=RS("CtNodeKind")%>">
<%  if (not memo_disable) then
		if RS("CatNameMemo") <> "" or RS("CatNameMemo")<> NULL then %>
			<Caption><![CDATA[<%=xCat%> (<%=RS("CatNameMemo")%>)]]></Caption>
<% 		else %>	
            <Caption><![CDATA[<%=xCat%>]]></Caption>	  
<% 		end if 
		else 
%>
			<Caption><![CDATA[<%=xCat%>]]></Caption>
<%
	end if
%>  
            <xCatShowOrder><%=RS("CatShowOrder")%></xCatShowOrder>
<%
        if RS("redirectURL")<> "" then
		xUrl = RS("redirectURL")
	elseif RS("CtUnitKind") ="2" then
		xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	elseif isNumeric(RS("iBaseDSD")) then
		xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	end if
%>   
          <xURL newWindow="<% =RS("newWindow") %>"><%=deAmp(xURL)%></xURL>        
          </Menucat> 
          
        
<%		
          else
		xUrl = ""
		xNewWindow = ""
		if isNumeric(RS("iBaseDSD")) then
			xUrl = "List.asp?ctNode="&RS("CtNodeID") & "&CtUnit=" & RS("CtUnitID") & "&BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
		elseif RS("redirectURL")<> "" then
			xUrl = RS("redirectURL")
			if RS("newWindow")="Y"  then	xNewWindow = " target=""_nwMof"""
		end if
%>
                
<% if xUrl<>"" then %>
<%
        if RS("redirectURL")<> "" then
		xUrl = RS("redirectURL")
	elseif RS("CtUnitKind") ="2" then
		xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	elseif isNumeric(RS("iBaseDSD")) then
		xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	end if
%>        
          <Menucat iCuItem="<%=RS("CtNodeID")%>" dataparent="<%=RS("DataParent")%>" isUnit="<%=RS("CtNodeKind")%>"> 
<% 	 if (not memo_disable) then
		if (RS("CatNameMemo") <> "") or (RS("CatNameMemo")<> NULL) then %>
			<Caption><![CDATA[<%=RS("CatName")%> (<%=RS("CatNameMemo")%>)]]></Caption>
<% 		else %>	
            <Caption><![CDATA[<%=RS("CatName")%>]]></Caption>	  
<% 		end if 
	else
%>
		<Caption><![CDATA[<%=RS("CatName")%>]]></Caption>
<%
	end if %>
            <xCat><%=RS("xCat")%></xCat>
            <xURL newWindow="<% =RS("newWindow") %>"><%=deAmp(xURL)%></xURL>  
            <xCatShowOrder><%=RS("CatShowOrder")%></xCatShowOrder>
          </Menucat> 		  
<% else %>
	 <Menucat iCuItem="<%=RS("CtNodeID")%>" dataparent="<%=RS("DataParent")%>" isUnit="<%=RS("CtNodeKind")%>">
<% 	if (not memo_disable) then
		if RS("CatNameMemo") <>"" or RS("CatNameMemo")<> NULL then %>
			<Caption><![CDATA[<%=RS("CatName")%> (<%=RS("CatNameMemo")%>)]]></Caption>
<%		 else %>	
            <Caption><![CDATA[<%=RS("CatName")%>]]></Caption>	  
<% 		end if
	else
%>
		<Caption><![CDATA[<%=RS("CatName")%>]]></Caption>
<%		
	end if %> 
            <xCat><%=RS("xCat")%></xCat>
            <xCatShowOrder><%=RS("CatShowOrder")%></xCatShowOrder>
          </Menucat> 
<% end if %>
               
<%		
           end if
		RS.moveNext
	wend
	
   rs0.Movenext
   loop	
%>	

        </Sitemap> 
        
        	<Sitemap2 myTreeNode="<%=myTreeNode%>">
<%

	SQLCom = "select * from CatTreeRoot Where CtRootID = "& xSitemapID	
	set RS = conn.execute(SqlCom)

        if not rs.eof then
	FtypeName = RS("CtRootName")

   sql="select distinct datalevel from CatTreeNode where CtRootID = " & xSitemapID
   set rs0=conn.Execute(sql)
   do while not rs0.eof   
    if rs0(0)="1" then
        SqlCom = "SELECT a.CatName AS xCat,a.*,b.redirectURL, b.newWindow,b.iBaseDSD,b.CtUnitKind "_
                & " FROM CatTreeNode a,CtUnit b"_  
                & " WHERE  b.ctUnitID=*a.CtUnitID and a.inUse='Y' AND  a.DataLevel=" & rs0(0) & " AND  a.CtRootID = " & xSitemapID  _
                & " Order by  a.CatShowOrder"
   else
 	SqlCom = "SELECT b.CatName AS xCat, a.*, c.redirectURL, c.newWindow, c.iBaseDSD , c.CtUnitKind " _
		& " FROM CatTreeNode AS a JOIN CatTreeNode AS b ON b.CtNodeID=a.DataParent" _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.inUse='Y' AND a.DataLevel=" & rs0(0) & " AND a.CtRootID = " & xSitemapID  _
		& " Order by b.CatShowOrder, a.CatShowOrder"
    end if		
	'response.write SqlCom & "<HR>"

	set RS = conn.execute(SqlCom)
	xCat = ""
	while not RS.eof
	  if rs0(0)="1" then
		xCat = RS("xCat")
%>
           <Menucat iCuItem="<%=RS("CtNodeID")%>" dataparent="<%=RS("DataParent")%>" isUnit="<%=RS("CtNodeKind")%>">
<% 	if (not memo_disable) then
		if RS("CatNameMemo") <>"" or RS("CatNameMemo")<> NULL and (not memo_disable) then %>
			<Caption><![CDATA[<%=xCat%> (<%=RS("CatNameMemo")%>)]]></Caption>
<% 		else %>	
            <Caption><![CDATA[<%=xCat%>]]></Caption>	  
<% 		end if
	else
%>
		<Caption><![CDATA[<%=xCat%>]]></Caption>
<%
	end if %>  
            <xCatShowOrder><%=RS("CatShowOrder")%></xCatShowOrder>
<%
        if RS("redirectURL")<> "" then
		xUrl = RS("redirectURL")
	elseif RS("CtUnitKind") ="2" then
		xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	elseif isNumeric(RS("iBaseDSD")) then
		xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")	
	end if
%>   
          <xURL newWindow="<% =RS("newWindow") %>"><%=deAmp(xURL)%></xURL>        
          </Menucat> 
          
        
<%		
          else
		xUrl = ""
		xNewWindow = ""
		if isNumeric(RS("iBaseDSD")) then
			xUrl = "List.asp?ctNode="&RS("CtNodeID") & "&CtUnit=" & RS("CtUnitID") & "&BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
		elseif RS("redirectURL")<> "" then
			xUrl = RS("redirectURL")
			if RS("newWindow")="Y"  then	xNewWindow = " target=""_nwMof"""
		end if
%>
                
<% if xUrl<>"" then %>
<%
        if RS("redirectURL")<> "" then
		xUrl = RS("redirectURL")
	elseif RS("CtUnitKind") ="2" then
		xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	elseif isNumeric(RS("iBaseDSD")) then
		xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
	end if
%>        
          <Menucat iCuItem="<%=RS("CtNodeID")%>" dataparent="<%=RS("DataParent")%>" isUnit="<%=RS("CtNodeKind")%>">
<% 	if (not memo_disable) then
		if RS("CatNameMemo") <>"" or RS("CatNameMemo")<> NULL then %>
			<Caption><![CDATA[<%=RS("CatName")%> (<%=RS("CatNameMemo")%>)]]></Caption>
<% 		else %>	
            <Caption><![CDATA[<%=RS("CatName")%>]]></Caption>	  
<% 		end if
	else
%>
		<Caption><![CDATA[<%=RS("CatName")%>]]></Caption>
<%
	end if %> 			
            <xCat><%=RS("xCat")%></xCat>
            <xURL newWindow="<% =RS("newWindow") %>"><%=deAmp(xURL)%></xURL>  
            <xCatShowOrder><%=RS("CatShowOrder")%></xCatShowOrder>
          </Menucat>  
<% else %>
	 <Menucat iCuItem="<%=RS("CtNodeID")%>" dataparent="<%=RS("DataParent")%>" isUnit="<%=RS("CtNodeKind")%>">
<% 	if (not memo_disable) then
		if RS("CatNameMemo") <>"" or RS("CatNameMemo")<> NULL then %>
			<Caption><![CDATA[<%=RS("CatName")%> (<%=RS("CatNameMemo")%>)]]></Caption>
<% 		else %>	
            <Caption><![CDATA[<%=RS("CatName")%>]]></Caption>	  
<% 		end if
	else
%>
		<Caption><![CDATA[<%=RS("CatName")%>]]></Caption>	
<%
	end if %>   
            <xCat><%=RS("xCat")%></xCat>
            <xCatShowOrder><%=RS("CatShowOrder")%></xCatShowOrder>
          </Menucat> 
<% end if %>
               
<%		
           end if
		RS.moveNext
	wend
	
   rs0.Movenext
   loop	
   
   end if
%>	

        </Sitemap2> 


<%	
	for each xDataSet in refModel.selectNodes("DataSet[ContentData='Y']")
		processXDataSet
	next
 
  
function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function

'--------------頁尾維護單位---------------------
footer_sql = "select footer_dept, footer_dept_url from cattreeroot left join nodeinfo on cattreeroot.ctrootid = nodeinfo.ctrootid where cattreeroot.pvXdmp ='"& request("mp") &"'"
set footer_rs = conn.Execute(footer_sql)
Response.Write "<footer_dept>" & footer_rs("footer_dept") & "</footer_dept>"
Response.Write "<footer_dept_url>" & footer_rs("footer_dept_url") & "</footer_dept_url>"

'--------------頁尾維護單位-end---------------
%>

<!--#include file="x1Menus.inc" -->
</hpMain>

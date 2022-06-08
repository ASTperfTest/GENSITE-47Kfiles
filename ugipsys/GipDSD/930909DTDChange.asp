<%@ CodePage = 65001 %>
<%
'-----930910修改xBody datalen/inputlen長度
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

'----檢查CuDTGeneric.xBody的欄位型態是否為text,若否,結束
Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open session("ODBCDSN")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

goFlag = false
fSql = "sp_columns 'CuDTGeneric'"
set RSlist = conn.execute(fSql)
if not RSlist.EOF then
    while not RSlist.eof
	if RSlist("COLUMN_NAME") = "xBody" then
		if RSlist("TYPE_NAME") = "text" then goFlag = true
	end if
      	RSlist.moveNext
    wend
end if

if not goFlag then
	response.write "CuDTGeneric資料表的xBody欄位型態請改為text!"
	response.end
end if

'----修改DTD/DSD
Set fso = server.CreateObject("Scripting.FileSystemObject")
xPath = "/site/"+session("mySiteID")+"/GipDSD/"
Set fldr = fso.GetFolder(server.MapPath(xPath))
	
if fldr.Files.count > 0 then
    for each sf in fldr.Files
    	if (Left(sf.name,5)="CuDTx" or Left(sf.name,7)="CtUnitX") and Left(sf.name,12)<>"CtUnitXOrder" then
		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true            	
		LoadXML = server.MapPath(xPath+sf.name)            	
    		xv = htPageDom.load(LoadXML)
  		if htPageDom.parseError.reason <> "" then 
    			Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    			Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    			Response.End()
  		end if 
'----修改datalen/inputlen長度  		
	    	if nullText(htPageDom.selectSingleNode("//fieldList/field[fieldName='xBody']"))<>"" then	
			Set fieldNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='xBody']")			
  			fieldNode.selectSingleNode("dataLen").text = ""
  			fieldNode.selectSingleNode("inputLen").text = ""
			htPageDom.save(server.MapPath(xPath+sf.name))
			response.write sf.name+"<br>"
		end if

'----回存  		
   	end if
    next
end if

response.write "<br>DONE930903!"
%>

<%@ CodePage=65001 Language="VBScript"%>
<?xml version="1.0"  encoding="utf-8" ?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#Include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
'----新增----
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

	specID=request.querystring("specID")
	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	
	set oxsl = server.createObject("microsoft.XMLDOM")
	'----xmlSpec處理開始
	xv = oxml.load(server.mappath("xmlSpec/"&specID&"Add.xml"))
	if oxml.parseError.reason <> "" then 
		Response.Write("htPageDom parseError on line " &  oxml.parseError.line)
		Response.Write("<BR>Reason: " &  oxml.parseError.reason)
		Response.End()
	end if
	'----代碼tag更換
	for each fieldNode in oxml.selectNodes("//fieldList/field[refLookup!='']")
	    SQL="Select * from CodeMetaDef where codeID='" & nullText(fieldNode.selectSingleNode("refLookup")) & "'"
	    SET RSLK=conn.execute(SQL)  
	    if not RSLK.eof then
	    	SQLRSS="Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where 1=1"
	    	if not isNull(RSLK("CodeSrcFld")) then _
	    		SQLRSS = SQLRSS & " AND " & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'"
	    	SQLRSS = SQLRSS & " Order by " & RSLK("CodeSortFld") 
	    	Set RSS=conn.execute(SQLRSS)
	    	if not RSS.eof then
	    	    codeStr=""
		    while not RSS.eof
		    	codeStr=codeStr&"<code name="""&RSS(0)&""" value="""&RSS(1)&"""/>"
		    	RSS.movenext
		    wend
		    codeStr="<refLookup>"&codeStr&"</refLookup>"
		    set nxml0 = server.createObject("microsoft.XMLDOM")
		    nxml0.LoadXML(codeStr)
		    set newNode0 = nxml0.documentElement	
		    fieldNode.replaceChild newNode0,fieldNode.selectSingleNode("refLookup")	    
	        end if
	    end if
	next
	'----xmlSpec處理結束
	'----Load XSL
	oxsl.load(server.mappath("xslGip/PatternAdd.xsl"))
	oxsl.selectSingleNode("//xsl:template[@match='htPage']/html/body/form/@action").text="add_do_action.asp?specID="&specID
	'----轉換為HTML字串
	Response.ContentType = "text/HTML" 
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	outString = replace(outString,"&amp;","&")
	response.write replace(outString,"&amp;","&")		
%>


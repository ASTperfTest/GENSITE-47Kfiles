<%@  codepage="65001" %>
<!--#include virtual = "/inc/GetPediaTable.inc" -->
<%
		'小百科字辭及時解釋(取代字串)
		'2011/05/26 Modify by Max
		haveBody = false

		Function CheckInput(str,strType)
		Dim strTmp
			strTmp = ""
		If strType = "s" Then
			strTmp = Replace(Trim(str),"@#","@#@#")
			strTmp = Replace(Trim(strTmp),";","")
			strTmp = Replace(Trim(strTmp),"'","")
			strTmp = Replace(Trim(strTmp),"--","")
			strTmp = Replace(Trim(strTmp),"/*","")
			strTmp = Replace(Trim(strTmp),"*/","")
			strTmp = Replace(Trim(strTmp),"*","")
			strTmp = Replace(Trim(strTmp),"/","")
			strTmp = Replace(Trim(strTmp),"<","")
			strTmp = Replace(Trim(strTmp),">","")

		ElseIf strType="i" Then
		If isNumeric(str)=False Then str="0"
			strTmp = str
		Else
			strTmp = str
		End If
				
		CheckInput = strTmp
		
		End Function
		
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open session("ODBCDSN")
		
		Dim iCuItemID,mp,rowId
		
		iCuItemID=CheckInput(request("iCUItem"),"i")
		mp = CheckInput(request("mp"),"i")
		rowId=CheckInput(request("themeId"),"i")
		review=CheckInput(request("review"),"i")
		
	
		'建立小百科 table
		response.write GetPediaTable(iCuItemID,mp,rowId)
			
		%>
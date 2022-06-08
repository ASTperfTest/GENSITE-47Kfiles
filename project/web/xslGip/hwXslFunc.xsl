<?xml version="1.0" encoding="utf-8" ?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" 
xmlns:hyweb="urn:gip-hyweb-com" version="1.0">

<msxsl:script language="VBScript" implements-prefix="hyweb"><![CDATA[
Function engFullDate(nodeList) 
	dim eMonth(12)
	eMonth(1) = "January"
	eMonth(2) = "February"
	eMonth(3) = "March"
	eMonth(4) = "April"
	eMonth(5) = "May"
	eMonth(6) = "June"
	eMonth(7) = "July"
	eMonth(8) = "August"
	eMonth(9) = "September"
	eMonth(10) = "October"
	eMonth(11) = "November"
	eMonth(12) = "December"
  Set myNode=nodeList(0)
   xDate=cdate(myNode.nodeTypedvalue)
   rStr = eMonth(month(xDate)) & " " & day(xDate) & ", " & year(xDate)
   engFullDate = rStr
End Function

Function dayOfWeek(nodeList)
  Set myNode=nodeList(0)
   xDate=cdate(myNode.nodeTypedvalue)
   rStr = WeekdayName(Weekday(xDate), ture)
   dayOfWeek = rStr
  
End Function

Function edayOfWeek(nodeList)
	dim eWeek(7)
	eWeek(1) = "(日)"
	eWeek(2) = "(一)"
	eWeek(3) = "(二)"
	eWeek(4) = "(三)"
	eWeek(5) = "(四)"
	eWeek(6) = "(五)"
	eWeek(7) = "(六)"
  Set myNode=nodeList(0)
   xDate=cdate(myNode.nodeTypedvalue)
   rStr = eWeek(Weekday(xDate))
   edayOfWeek = rStr
End Function

Function engDate(nodeList) 
	dim eMonth(12)
	eMonth(1) = "Jan."
	eMonth(2) = "Feb."
	eMonth(3) = "Mar."
	eMonth(4) = "Apr."
	eMonth(5) = "May "
	eMonth(6) = "Jun."
	eMonth(7) = "Jul."
	eMonth(8) = "Aug."
	eMonth(9) = "Sep."
	eMonth(10) = "Oct."
	eMonth(11) = "Nov."
	eMonth(12) = "Dec."
  Set myNode=nodeList(0)
   xDate=cdate(myNode.nodeTypedvalue)
   rStr = eMonth(month(xDate)) & " " & day(xDate) & ", " & year(xDate)
   engDate = rStr
End Function

Function twDate(nodeList)
  Set myNode=nodeList(0)
   xDate=cdate(myNode.nodeTypedvalue)
   rStr = "民國" & year(xDate)-1911 & "年" & month(xDate)& "月" & day(xDate) & "日"
   twDate = rStr
End Function

Function chDate(nodeList)
  Set myNode=nodeList(0)
	if isDate(myNode.nodeTypedvalue) then
	   xDate=cdate(myNode.nodeTypedvalue)
	   rStr = year(xDate)-1911 & "/" & month(xDate)& "/" & day(xDate)
	else
		rStr = ""
	end if
   chDate = rStr
End Function

]]></msxsl:script>
</xsl:stylesheet>


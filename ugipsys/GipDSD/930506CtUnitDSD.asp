<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題單元管理"
HTProgFunc="欄位選取"
HTProgCode="GE1T11"
HTProgPrefix="CtUnitDSD" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<%
function checkboxYN(xNode, xFun, xList, xName)
	xStr = ""
  	on error resume next
  	xNodestr = ""
  	xNodestr = xNode.text
  	xListstr = ""
  	xListstr = xList.text 
  	xNamestr = ""
  	xNamestr = xName.text 
  	if xNamestr="sTitle" and xFun = "show" then	
  		xStr = " disabled "
  	elseif xNodestr = "N" and xFun = "form" then
  		xStr = " checked disabled "
  	elseif xListstr <> "" then
  		xStr = " checked "  	
  	end if
  	checkboxYN = xStr
end function

'----1.1依ctUnitID檢查/GipDSD/xmlSpec下是否存在CtUnitX???.xml(???=ctUnitID)
'------1.1.1若不存在, GipDSD/CuDTx???(???=baseDSD)
'----2.1開XMLDOM物件, load xml
'----3.1PSetListEdit方式批次編修

'----1.1xmlSpec檔案檢查
Set fso = server.CreateObject("Scripting.FileSystemObject")
LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX"+request("ctUnit")+".xml")
if not fso.FileExists(LoadXML) then
	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & request("baseDSD") & ".xml")
end if
'----2.1開XMLDOM物件, load .xml
set oxml = server.createObject("microsoft.XMLDOM")
oxml.async = false
oxml.setProperty "ServerHTTPRequest", true
xv = oxml.load(LoadXML)
if oxml.parseError.reason <> "" then
	Response.Write("XML parseError on line " &  oxml.parseError.line)
	Response.Write("<BR>Reason: " &  oxml.parseError.reason)
	Response.End()
end if 
'----轉換順序 
set oxsl = server.createObject("microsoft.XMLDOM")
oxsl.async = false
xv = oxsl.load(server.mappath("/GipDSD/xmlspec/CtUnitXOrder.xsl"))
set nxml0 = server.createObject("microsoft.XMLDOM")
nxml0.LoadXML(oxml.transformNode(oxsl))
set allModel = nxml0.selectNodes("fieldList/field[inputType!='hidden']")

%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
	<td width="50%" class="FormLink">
		<%if (HTProgRight and 4)=4 then%>
			<A href="javascript:window.history.back();" title="回主題單元編修介面">回前頁</A>
		<%end if%>
	</td>
	<td width="50%" class="FormLink" align="right">
   		 <input type=button class=cbutton value="編修存檔" id=button1>
        </td>   
    </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
 <Form name=reg method="POST" action="CtUnitDSDUpdate.asp?ctUnit=<%=request("ctUnit")%>&baseDSD=<%=request("baseDSD")%>">
 <input type=hidden name="submittask" value="">
     <TABLE width=100% cellspacing="1" cellpadding="0" class=bluetable>                   
     <tr align=left class="lightbluetable">    
	<td class=eTableLable rowspan=2 align=center>顯示<br>順序</td>
	<td class=eTableLable rowspan=2 align=center width=20%>欄位原有名稱</td>	
	<td class=eTableLable rowspan=2 align=center width=20%>欄位顯示名稱</td>
	<td class=eTableLable colspan=3 align=center>前台是否顯示?</td>
	<td class=eTableLable colspan=3 align=center>後台是否顯示?</td>
    </tr>
     <tr align=left class="lightbluetable">    
	<td class=eTableLable align=center>條列</td>
	<td class=eTableLable align=center>內容呈現</td>
	<td class=eTableLable align=center>查詢</td>
	<td class=eTableLable align=center>條列</td>
	<td class=eTableLable align=center>新增/編修</td>
	<td class=eTableLable align=center>查詢</td>	
    </tr>  
<%
for each param in allModel
	j=j+1
%>
<tr> 
	<TD class=eTableContent align=center><font size=2>
		<input type=hidden Name="htx_fieldName<%=j%>" value="<%=nullText(param.selectSingleNode("fieldName"))%>">	
		<input type=text Name="htx_fieldSeq<%=j%>" size="3" maxlength=3 value="<%=nullText(param.selectSingleNode("fieldSeq"))%>">
	</font></td>
	<TD class=eTableContent><font size=2><%=nullText(param.selectSingleNode("fieldLabelInit"))%>
	</font></td>
	<TD class=eTableContent><font size=2>
		<input type=text Name="htx_fieldLabel<%=j%>" size="40" value="<%=nullText(param.selectSingleNode("fieldLabel"))%>">
	</font></td>
	<TD class=eTableContent align=center><font size=2>
		<input type=checkbox name=ckboxshowListClient<%=j%> <%=checkboxYN(param.selectSingleNode("canNull"),"show",param.selectSingleNode("showListClient"),param.selectSingleNode("fieldName"))%>>
	</font></td>
	<TD class=eTableContent align=center><font size=2>
		<input type=checkbox name=ckboxformListClient<%=j%> <%=checkboxYN(param.selectSingleNode("canNull"),"form",param.selectSingleNode("formListClient"),param.selectSingleNode("fieldName"))%>>
	</font></td>
	<TD class=eTableContent align=center><font size=2>
		<input type=checkbox name=ckboxqueryListClient<%=j%> <%=checkboxYN(param.selectSingleNode("canNull"),"query",param.selectSingleNode("queryListClient"),param.selectSingleNode("fieldName"))%>>
	</font></td>
	<TD class=eTableContent align=center><font size=2>
		<input type=checkbox name=ckboxshowList<%=j%> <%=checkboxYN(param.selectSingleNode("canNull"),"show",param.selectSingleNode("showList"),param.selectSingleNode("fieldName"))%>>
	</font></td>	
	<TD class=eTableContent align=center><font size=2>
		<input type=checkbox name=ckboxformList<%=j%> <%=checkboxYN(param.selectSingleNode("canNull"),"form",param.selectSingleNode("formList"),param.selectSingleNode("fieldName"))%>>
	</font></td>
	<TD class=eTableContent align=center><font size=2>
		<input type=checkbox name=ckboxqueryList<%=j%> <%=checkboxYN(param.selectSingleNode("canNull"),"query",param.selectSingleNode("queryList"),param.selectSingleNode("fieldName"))%>>
	</font></td>					
    </tr>
<%
next
%>      
    </TABLE>    	                             
</table> 
</form>
</body>
</html>  
<script Language=VBScript>
sub button1_onClick
	 set xItem=document.all.item("reg")
  	 for i=0 to xItem.length-1  	      
	      	   if xItem.item(i).tagname = "INPUT" then
	                if Left(xItem.item(i).name,12)="htx_fieldSeq" then
		              xn=mid(xItem.item(i).name,13)  	           
	 	              if reg("htx_fieldSeq"&xn).value ="" then 
                                   alert "顯示順序欄位不能空白且須為數字(小於999)"
                                   reg("htx_fieldSeq"&xn).focus
                                   exit sub
                              end if
	 	              if not IsNumeric(reg("htx_fieldSeq"&xn).value) then 
                                   alert "顯示順序欄位不能空白且須為數字"
                                   reg("htx_fieldSeq"&xn).focus
                                   exit sub
                              end if                              
                        end if
	      	   end if 
	  next   
 	  reg.submit   
end sub
</script>
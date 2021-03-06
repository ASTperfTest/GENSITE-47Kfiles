<%@ CodePage = 65001 %>
<% Response.Expires = 0
'------ Modify History List (begin) ------
' 2006/1/6	92004/Chirs
'	1. CtNodeDND.asp created from /GipDSD/CtUnitDSD.asp
'
'------ Modify History List (begin) ------
HTProgCap="目錄樹管理"
HTProgFunc="欄位選取"
HTProgCode="GE1T21"
HTProgPrefix="CtUnitNode" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
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

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

	sqlCom = "SELECT htx.*, u.ibaseDsd " _
		& " ,(Select count(*) from CatTreeNode where dataParent=htx.ctNodeId) childCount" _
		& " FROM CatTreeNode AS htx JOIN CtUnit AS u ON u.ctUnitId=htx.ctUnitId" _
		& " JOIN BaseDsd AS b ON b.ibaseDsd=u.ibaseDsd" _
		& " WHERE htx.ctNodeId=" & pkStr(request.queryString("ctNodeId"),"")
	Set RSreg = Conn.execute(sqlcom)
	if RSreg.eof then
			showDoneBox("資料不存在！")
			response.end
	end if

	myNodeID = RSreg("ctNodeId")
	myUnitID = RSreg("ctUnitId")
	myBaseID = RSreg("ibaseDsd")
	
'----1.0依ctNodeID檢查/GipDSD/xmlSpec下是否存在CtNodeX???.xml(???=ctNodeID)
'----1.1依ctUnitID檢查/GipDSD/xmlSpec下是否存在CtUnitX???.xml(???=ctUnitID)
'------1.1.1若不存在, GipDSD/CuDTx???(???=baseDSD)
'----2.1開XMLDOM物件, load xml
'----3.1PSetListEdit方式批次編修

'----1.0xmlSpec檔案檢查
Set fso = server.CreateObject("Scripting.FileSystemObject")
LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtNodeX" & myNodeID & ".xml")
if not fso.FileExists(LoadXML) then
	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & myUnitID & ".xml")
	if not fso.FileExists(LoadXML) then
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & myBaseID & ".xml")
	end if
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
set allModelref = nxml0.selectNodes("fieldList/field[inputType!='hidden' and (inputType='refSelect' or inputType='refSelectOther' or inputType='refCheckbox' or inputType='refCheckboxOther' or inputType='refRadio' or inputType='refRadioOther')]")
'----orderby初始值
xshowClientSqlOrderBy=""
xshowClientSqlOrderByType=""
if nullText(oxml.selectSingleNode("//DataSchemaDef/showClientSqlOrderBy"))<>"" then
	OrderByStr=nullText(oxml.selectSingleNode("//DataSchemaDef/showClientSqlOrderBy"))
'response.write 	OrderByStr & "<br>"
	if instr(OrderByStr,"DESC")<>0 then
		xshowClientSqlOrderByType="DESC"
		OrderByStr=Left(OrderByStr,Len(OrderByStr)-5)
	end if
	xshowClientSqlOrderBy=mid(OrderByStr,10)
end if
'response.write xshowClientSqlOrderBy&"[]"&xshowClientSqlOrderByType 
xformClientCatRefLookup=""
if nullText(oxml.selectSingleNode("//DataSchemaDef/formClientCat"))<>"" then
	xformClientCatRefLookup=nullText(oxml.selectSingleNode("//DataSchemaDef/dsTable/fieldList/field[fieldName='"+nullText(oxml.selectSingleNode("//DataSchemaDef/formClientCat"))+"']/refLookup"))
end if
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
 <Form name=reg method="POST" action="CtNodeDNDUpdate.asp?ctNodeId=<%=myNodeId%>">
  <tr>
	<td class="FormName"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
	<td class="FormLink" align=right>
		<%if (HTProgRight and 4)=4 then%>
			<A href="javascript:window.history.back();" title="回節點編修介面">回前頁</A>
		<%end if%>
	</td>	    
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
	<td class="FormLink">
	    前台條列排序欄位:
	    <Select name="showClientSqlOrderBy" size=1>
		<option value="">請選擇</option>
<%		for each param in allModel%>
			<option value="<%=nullText(param.selectSingleNode("fieldName"))%>"><%=nullText(param.selectSingleNode("fieldLabel"))%></option>
<%		next%>
	    </select>
	    升/降羃:
	    <Select name="showClientSqlOrderByType" size=1>
		<option value="">升羃</option>
		<option value="DESC">降羃</option>
	    </select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;	    
	    分類欄位:
	    <Select name="formClientCat" size=1>
		<option value="">請選擇</option>
<%		for each param in allModelref%>
			<option value="<%=nullText(param.selectSingleNode("fieldName"))%>"><%=nullText(param.selectSingleNode("fieldLabel"))%></option>
<%		next%>
	    </select>
	    參照代碼:
	    <input Name="formClientCatRefLookup" value="" size=10>   
	</td>
	<td class="FormLink" align="right">
   		 <input type=button class=cbutton value="編修存檔" id=button1>
        </td>   
    </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
 <input type=hidden name="submittask" value="">
     <TABLE width=100% cellspacing="1" cellpadding="0" class=bluetable>                   
     <tr align=left class="lightbluetable">    
	<td class=eTableLable rowspan=2 align=center>顯示<br>順序</td>
	<td class=eTableLable rowspan=2 align=center width=10%>欄位原有名稱</td>	
	<td class=eTableLable rowspan=2 align=center width=10%>欄位顯示名稱</td>
	<td class=eTableLable rowspan=2 align=center width=12%>欄位說明</td>
	<td class=eTableLable colspan=3 align=center>前台是否顯示?</td>
	<td class=eTableLable colspan=3 align=center>後台是否顯示?</td>
    </tr>
     <tr align=left class="lightbluetable">    
	<td class=eTableLable align=center>條列
	    <Select name="showClientStyle" size=1>
			<%SQL="Select mcode,mvalue from CodeMain where codeMetaId='showClientStyle' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
	    </select>	
	</td>
	<td class=eTableLable align=center>內容呈現
	    <Select name="formClientStyle" size=1>
			<%SQL="Select mcode,mvalue from CodeMain where codeMetaId='formClientStyle' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
	    </select>	
	</td>
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
		<input type=text Name="htx_fieldLabel<%=j%>" value="<%=nullText(param.selectSingleNode("fieldLabel"))%>">
	</font></td>
	<TD class=eTableContent><font size=2>
		<input type=text Name="htx_fieldDesc<%=j%>" value="<%=nullText(param.selectSingleNode("fieldDesc"))%>">
	</font></td>
	<TD class=eTableContent align=center><font size=2>
		<input type=checkbox name=ckboxshowListClient<%=j%> <%=checkboxYN(param.selectSingleNode("canNull"),"show",param.selectSingleNode("showListClient"),param.selectSingleNode("fieldName"))%>>
	</font></td>
	<TD class=eTableContent align=center><font size=2>
		<input type=checkbox name=ckboxformListClient<%=j%> <%=checkboxYN(param.selectSingleNode("canNull"),"formClient",param.selectSingleNode("formListClient"),param.selectSingleNode("fieldName"))%>>
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
</form>
</table> 
</body>
</html>  
<script Language=VBScript>
reg.showClientStyle.value="<%=nullText(oxml.selectSingleNode("//DataSchemaDef/showClientStyle"))%>"
reg.formClientStyle.value="<%=nullText(oxml.selectSingleNode("//DataSchemaDef/formClientStyle"))%>"
reg.showClientSqlOrderBy.value="<%=xshowClientSqlOrderBy%>"
reg.showClientSqlOrderByType.value="<%=xshowClientSqlOrderByType%>"
reg.formClientCat.value="<%=nullText(oxml.selectSingleNode("//DataSchemaDef/formClientCat"))%>"
reg.formClientCatRefLookup.value="<%=xformClientCatRefLookup%>"
redim refArray(1,0)
<%	for each param in allModelref
		i=i+1
%>
	redim preserve refArray(1,<%=i%>)
	refArray(0,<%=i%>)="<%=nullText(param.selectSingleNode("fieldName"))%>"
	refArray(1,<%=i%>)="<%=nullText(param.selectSingleNode("refLookup"))%>"
<%	next%>	

sub formClientCat_onChange()
	if reg.formClientCat.value="" then 
		reg.formClientCatRefLookup.value = ""
		exit sub
	end if
	for k=0 to ubound(refArray,2)
		if refArray(0,k)=reg.formClientCat.value then
			reg.formClientCatRefLookup.value=refArray(1,k)
		end if
	next
end sub

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
	  if reg.formClientCat.value<>"" and reg.formClientCatRefLookup.value="" then
	  	alert "選擇分類欄位後, 必須輸入參照代碼!"
	  	exit sub
	  end if
 	  reg.submit   
end sub
</script>

<%Sub showDoneBox(lMsg) 
	mpKey = "&phase=edit&CtNodeID="&request("ctNodeId")
	if mpKey<>"" then  mpKey = mid(mpKey,2)
	doneURI= "CtUnitNodeEdit.asp"
	if doneURI = "" then	doneURI = HTprogPrefix & "List.asp"
%>
    <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <meta name="GENERATOR" content="Hyweb Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>編修表單</title>
     </head>
     <body>
     <script language=vbs>
       	alert("<%=lMsg%>")
	    document.location.href="<%=doneURI%>?<%=mpKey%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
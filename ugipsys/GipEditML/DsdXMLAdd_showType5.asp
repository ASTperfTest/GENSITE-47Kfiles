<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="單元資料維護"
HTProgFunc="新增"
HTUploadPath=session("Public")+"data/"
HTProgCode="GC1AP1"
HTProgPrefix="DsdXML" %>
<!--#include virtual = "/inc/server.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
sub showType5AddProcessInit(param)

end sub

function nullText(xNode)
    on error resume next
    xstr = ""
    xstr = xNode.text
    nullText = xStr
end function

Function Send_Email (S_Email,R_Email,Re_Sbj,Re_Body)

	Set objNewMail = CreateObject("CDONTS.NewMail") 
	objNewMail.MailFormat = 0
	objNewMail.BodyFormat = 0 
	call objNewMail.Send(S_Email,R_Email,Re_Sbj,Re_Body)

	Set objNewMail = Nothing
End Function	

Dim pKey
Dim RSreg
Dim formFunction
Dim Sql, SqlValue
Dim xNewIdentity
Dim orgInputType
taskLable="新增" & HTProgCap
	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"
 
apath=server.mappath(HTUploadPath) & "\"
set xup = Server.CreateObject("UpDownExpress.FileUpload")
xup.Open 
function xUpForm(xvar)
on error resume next
	xStr = ""
	arrVal = xup.MultiVal(xvar)
	for i = 1 to Ubound(arrVal)
		xStr = xStr & arrVal(i) & ", "
'		Response.Write arrVal(i) & "<br>" & Chr(13)
	next 
	if xStr = "" then
		xStr = xup(xvar)
		xUpForm = xStr
	else
		xUpForm = left(xStr, len(xStr)-2)
	end if
end function

'function xUpForm(xvar)
'	xUpForm = request.form(xvar)
'end function

set htPageDom = session("codeXMLSpec")
set refModel = htPageDom.selectSingleNode("//dsTable")
set allModel = htPageDom.selectSingleNode("//DataSchemaDef")
    set iDeptNode = htPageDom.selectSingleNode("//fieldList/field[fieldName='iDept']")
    iDeptNode.selectSingleNode("inputType").text = "refSelect"  
    iDeptNode.selectSingleNode("refLookup").text = "refDept"  


	'----被引用資料區	
	SQLG="Select CDG.iBaseDSD,CDG.iCTUnit,B.sBaseTableName from CuDTGeneric CDG Left Join BaseDSD B on CDG.iBaseDSD=B.iBaseDSD " & _
		"where iCUItem=" & request.querystring("iCuItem")
	set RSC=conn.execute(SQLG)
	if isNull(RSC("sBaseTableName")) then
		xBaseTableName = "CuDTx" & RSC("iBaseDSD")
	else
		xBaseTableName = RSC("sBaseTableName")
	end if		
	sqlCom = "SELECT htx.*, ghtx.*, xrefNFileName.oldFileName AS fxr_fileDownLoad FROM " & xBaseTableName _
		& " AS htx JOIN CuDTGeneric AS ghtx ON ghtx.iCUItem=htx.giCuItem "_
		& " LEFT JOIN imageFile AS xrefNFileName ON xrefNFileName.newFileName = ghtx.fileDownLoad " _
		& " WHERE ghtx.iCUItem=" & pkStr(request.queryString("iCuItem"),"")
	Set RSreg = Conn.execute(sqlcom)	
	set htPageDomRef = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDomRef.async = false
	htPageDomRef.setProperty("ServerHTTPRequest") = true		
    	'----找出對應的CtUnitX???? xmlSpec檔案(若找不到則抓default)
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filePath = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(RSC("iCTUnit")) & ".xml") 	
    	if fso.FileExists(filePath) then
    		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & cStr(RSC("iCTUnit")) & ".xml")
    	else
    		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & RSC("iBaseDSD") & ".xml")
    	end if 
	xv = htPageDomRef.load(LoadXML)	
	set allModelRef = htPageDomRef.selectSingleNode("//DataSchemaDef")	
	'----新增資料區
    	'----Load XSL樣板
    	set oxsl2 = server.createObject("microsoft.XMLDOM")
   	oxsl2.async = false
   	xv = oxsl2.load(server.mappath("/GipDSD/xmlspec/CtUnitXOrder.xsl"))    	
	set nxml2 = server.createObject("microsoft.XMLDOM")
	nxml2.LoadXML(htPageDom.transformNode(oxsl2))
	set allModel2 = nxml2.documentElement  
	for each param in allModel2.selectNodes("//fieldList/field[showTypeStr='3']") 
		set romoveNode=allModel2.selectSingleNode("field[fieldName='"+param.selectSingleNode("fieldName").text+"']")
		allModel2.removeChild romoveNode
	next
  	xShowType=request.querystring("showType")

'response.write "<XMP>"+allModel2.xml+"</XMP>"
'response.end	

if xUpForm("submitTask") = "ADD" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		showErrBox()
	else
		doUpdateDB()
		showDoneBox()
	end if

else
	showHTMLHead()
	formFunction = "add"
	showForm()
	initForm()
	showHTMLTail()
end if
function qqRS(fldName)
	if request("submitTask")="" then
		xValue = RSreg(fldname)
	else
		xValue = ""
		if request("htx_"&fldName) <> "" then
			xValue = request("htx_"&fldName)
		end if
	end if
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
'		xp = instr(xqqRS,vbCRLF&vbCRLF)
'		while xp > 0
'			xqqRS = left(xqqRS,xp-1) & mid(xqqRS,xp+4)
'			xp = instr(xqqRS,vbCRLF&vbCRLF)
'		wend
		xqqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
		xqqRS = replace(xqqRS,chr(13),"")
		qqRS = replace(xqqRS,chr(10),"")
	end if
end function
%>
<% sub initForm() %>
<script language=vbs>
cvbCRLF = vbCRLF
cTabchar = chr(9)

function ncTabChar(n)
	if cTabchar = "" then
		ncTabChar = ""
	else
		ncTabChar = String(n,cTabchar)
	end if
end function

sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm	'---- 新增時的表單預設值放在這裡
	reg.showType.value="<%=xShowType%>"
	reg.refID.value="<%=request.querystring("iCuItem")%>"
<%
	for each xparam in allModel2.selectNodes("//fieldList/field[formList!='']") 
	    	if not (nullText(xparam.selectSingleNode("fieldName"))="fCTUPublic" and session("CheckYN")="Y") then
			orgInputType = xparam.selectSingleNode("inputType").text
			if nullText(xparam.selectSingleNode("fieldName"))="iDept" or nullText(xparam.selectSingleNode("fieldName"))="fCTUPublic" then
				xparam.selectSingleNode("inputType").text = "showType5"
			else
				xparam.selectSingleNode("inputType").text = "textbox"
			end if
			AddProcessInit xparam
			xparam.selectSingleNode("inputType").text = orgInputType		    	
	    	end if
	next	
	for each param in allModel2.selectNodes(("//fieldList/field[formList!='' and fieldName!='iCUItem' and fieldName!='iBaseDSD' and fieldName!='iCtUnit' and fieldName!='iDept' and fieldName!='fCTUPublic' and fieldName!='iEditor' and fieldName!='dEditDate' and fieldName!='Created_Date' and fieldName!='showType' and fieldName!='refID']"))
	    if nullText(allModelRef.selectSingleNode("//fieldList/field[fieldName='"&param.selectSingleNode("fieldName").text&"']"))<>"" then
	    	if nullText(param.selectSingleNode("fieldRefEditYN"))="N" then
			orgInputType = param.selectSingleNode("inputType").text
			param.selectSingleNode("inputType").text = "showType5"
			EditProcessInit param
			param.selectSingleNode("inputType").text = orgInputType			    	
	    	else
			EditProcessInit param
		end if
	    end if
	next
%>
end sub

sub window_onLoad
	clientInitForm
end sub


</script>	
<% end sub '---- initForm() ----%>

<% sub showForm() 	'===================== Client Side Validation Put HERE =========== %>
<script language="vbscript">

Sub formAddSubmit()	
    
	nMsg = "請務必填寫「{0}」，不得為空白！"
	lMsg = "「{0}」欄位長度最多為{1}！"
	dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
	iMsg = "「{0}」欄位應為數值！"
	pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"
<%
	for each param in allModel2.selectNodes("//fieldList/field[formList!='' and inputType!='hidden']") 
	    if not (nullText(param.selectSingleNode("fieldName"))="fCTUPublic" and session("CheckYN")="Y") then	
		processValid param
	    end if
	next
%>
  
	reg.submitTask.value = "ADD"
  	reg.Submit
End Sub

</script>

<!--#include file="DsdXMLForm_showType5.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>

<body>
<div id="FuncName">
	<h1><%=Title%>(引用資料式--參照新增)</h1>
	<div id="Nav">
	       <%if (HTProgRight and 4)=4 then%>
	       <a href="DsdXMLAdd.asp">回新增</a>	    
	       <%end if%>
	       <a href="Javascript:window.history.back();">回前頁</a>	    
	</div>
	<div id="ClearFloat"></div>
</div>
<% end sub '--- showHTMLHead() ------%>


<% sub ShowHTMLTail() %>     
</body>
</html>
<% end sub '--- showHTMLTail() ------%>


<% sub showErrBox() %>
	<script language=VBScript>
  			alert "<%=errMsg%>"
  			window.history.back
	</script>
<%
end sub '---- showErrBox() ----

sub checkDBValid()	'===================== Server Side Validation Put HERE =================

'---- 後端資料驗證檢查程式碼放在這裡	，如下例，有錯時設 errMsg="xxx" 並 exit sub ------
'	SQL = "Select * From Client Where ClientID = '"& Request("pfx_ClientID") &"'"
'	set RSvalidate = conn.Execute(SQL)
'
'	if not RSvalidate.EOF Then
'		errMsg = "「客戶編號」重複!!請重新建立客戶編號!"
'		exit sub
'	end if

end sub '---- checkDBValid() ----

function d6date(dt)     '轉成民國年  999/99/99 給資料型態為SmallDateTime 使用
	if Len(dt)=0 or isnull(dt) then
	     d6date=""
	else
                      xy=right("000"+ cstr((year(dt)-1911)),2)     '補零
                      xm=right("00"+ cstr(month(dt)),2)
                      xd=right("00"+ cstr(day(dt)),2)
                      d6date=xy &  xm &  xd
                  end if
end function

sub doUpdateDB()
	sql = "INSERT INTO  CuDTGeneric(showType,"
	sqlValue = ") VALUES('"+xUpForm("showType")+"'," 
	if xUpForm("showType")="4" or xUpForm("showType")="5" then
		sql = sql & "refID,"
		sqlValue = sqlValue & xUpForm("refID") & ","
	end if
	for each param in allModel.selectNodes("dsTable[tableName='CuDTGeneric']/fieldList/field[formList!='']") 
		if nullText(param.selectSingleNode("fieldName")) = "xImportant" _
				 AND xUpForm("xxCheckImportant")="Y" then
			sql = sql & "xImportant,"
			sqlValue = sqlValue & pkStr(d6date(date()),",")
		elseif not ((nullText(param.selectSingleNode("fieldName"))="fCTUPublic" and session("CheckYN")="Y") or (nullText(param.selectSingleNode("identity"))="Y")) then	
			processInsert param
		end if
	next

	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
	sql = "set nocount on;"&sql&"; select @@IDENTITY as NewID"
'	response.write sql
'	response.end
	set RSx = conn.Execute(SQL)
	xNewIdentity = RSx(0)
'	response.write sql & "<HR>"
	'----不同新增方式處理slave資料	
	    sql = "INSERT INTO  " & nullText(refModel.selectSingleNode("tableName")) & "(giCuItem,"
	    sqlValue = ") VALUES(" & dfn(xNewIdentity)
	    for each param in refModel.selectNodes("fieldList/field[formList!='']") 
		processInsert param
	    next
	    sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
	    conn.Execute(SQL) 
	'----Email處理
	if session("CheckYN")="Y" then
		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\sysPara.xml"
		xv = htPageDom.load(LoadXML)
  		if htPageDom.parseError.reason <> "" then 
	    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    			Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    			Response.End()
  		end if
		set emailNode=htPageDom.selectSingleNode("SystemParameter/DsdXMLEmail")
		xEmail=nullText(emailNode)
		xEmailStr=nullText(emailNode.selectSingleNode("@xdesc"))
		fSql="Select IU.UserName,IU.Email " & _
			" from CuDTGeneric C " & _
			" Left Join CatTreeNode CTN ON C.iCtUnit=CTN.CtUnitID AND CTN.CtRootID=" & session("ItemID") & _
			" Left Join CtUserSet2 CUS2 ON CTN.CtNodeID=CUS2.CtNodeID AND CUS2.UserID=N'"&session("userID")&"' " & _
			" Left Join CtUnit CT On C.iCtUnit=CT.CtUnitID " & _
			" Left Join InfoUser IU ON CUS2.userID=IU.userID AND C.iDept=IU.deptID " & _
			" where C.iCuitem="&xNewIdentity
		set rs1=conn.execute(fSql)
		if not rs1.EOF then
    		    while not rs1.EOF
        	        S_Email=""""+xEmailStr+""" <"+xEmail+">"
        		R_Email=rs1(1)
            
        		Email_body="【 " & rs1(0) & " 】小姐先生 您好:" & "<br>" & "<br>" & _
                            "   現有新的待審上稿資料, 請至[後台管理網站/資料審稿作業中]審核!"& "<br><br>" & _
                            "謝謝您!"& "<br>" & "<br>" & _
                            xEmailStr & "<br>" 	
			if not isNull(rs1(1)) then                                
        			Call Send_Email(S_Email,R_Email, "上稿資料審稿通知" ,Email_body)  
		        end if 
     
		     	rs1.movenext  
		    wend
		end if	

	end if	
	'----關鍵字詞處理
	if xUpForm("htx_xKeyword")<>"" then
	    redim iArray(1,0)
	    xStr=""
	    xReturnValue=""
	    SQLInsert=""
	    xKeywordArray=split(xUpForm("htx_xKeyword"),",")
	    weightsum=0
	    for i=0 to ubound(xKeywordArray)
	    	redim preserve iArray(1,i)
		'----分開字詞與權重符號
		xPos=instr(xKeywordArray(i),"*")
		if xPos<>0 then
			xStr=Left(trim(xKeywordArray(i)),xPos-1)
			iArray(0,i)=xStr
			iArray(1,i)=mid(xKeywordArray(i),xPos+1)
		else
			xStr=trim(xKeywordArray(i))
			iArray(0,i)=xStr
			iArray(1,i)=1		
		end if	
		weightsum=weightsum+iArray(1,i)
	    next   
	    '----串SQL字串 
	    for k=0 to ubound(iArray,2)
	    	SQLInsert=SQLInsert+"Insert Into CuDTKeyword values("+dfn(xNewIdentity)+"'"+iArray(0,k)+"',"+cStr(round(iArray(1,k)*100/weightsum))+");"
	    next
	    if SQLInsert<>"" then conn.execute(SQLInsert)
	end if
end sub '---- doUpdateDB() ----

%>

<% sub showDoneBox() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增程式</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	alert("新增完成！")
<%
	nextURL = HTprogPrefix & "List.asp"
	if xUpForm("nextTask") = "CuCheckList.asp" then
%>
		window.open "<%=session("myWWWSiteURL")%>/ContentOnly.asp?CuItem=<%=xNewIdentity%>&ItemID=<%=session("ItemID")%>&userID=<%=session("UserID")%>"
<%		
	elseif xUpForm("nextTask") <> "" then
		nextURL = xUpForm("nextTask") & "?iCuItem=" & xNewIdentity
	end if
%>
	document.location.href = "<%=nextURL%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>

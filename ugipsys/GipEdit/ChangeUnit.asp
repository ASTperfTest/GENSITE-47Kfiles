<%@ CodePage = 65001 %>
<% Response.Expires = 0
response.charset="utf-8"
'response.charset="big5"
HTProgCap=""
HTProgFunc="變更主題單元"
HTUploadPath="/"
HTProgCode="GC1AP1"
HTProgPrefix="DsDXML" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!--#include virtual = "/inc/hyftdGIP.inc" -->

<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function
%>
<%
taskLable="查詢" & HTProgCap
	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

  	set htPageDom = session("codeXMLSpec")
  	set allModel2 = session("codeXMLSpec2")  	  	
  	set refModel = htPageDom.selectSingleNode("//dsTable")
  	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")

	showHTMLHead()
	formFunction = "query"
	showForm()
	initForm()
	showHTMLTail()
%>

<% sub initForm() %>
<script language=vbs>
sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm

    end sub

    sub window_onLoad  
	     clientInitForm
    end sub
</script>	
<%end sub '---- initForm() ----%>

<%Sub showForm() %>

<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<object data="/inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
<CENTER>
<TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">

</TABLE>
</CENTER>
</form>     

<SCRIPT LANGUAGE="vbs">
    sub setOtherCheckbox(xname,otherValue)
    	reg.all(xname).item(reg.all(xname).length-1).value = otherValue
    end sub
</SCRIPT>
	
<%end sub '--- showForm() ------%>
<script language="javascript">
function unitList_onchange()
{
	window.location = "ChangeUnit.asp?ispostback=1&changeUnit=" + document.form1.unitlist.value;
}

function ispublic_onchange()
{
	window.location = "ChangeUnit.asp?ispostback=1&ispublic=" + document.form1.ispublic.value;
}
function datatype_onchange()
{
	window.location = "ChangeUnit.asp?ispostback=1&topCat=" + document.form1.datatype.value;
}
</script>
<%sub showHTMLHead() %>
    <html> 
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta name="GENERATOR" content="Hometown Code Generator 1.0">
    <meta name="ProgId" content="<%=HTprogPrefix%>Query.asp">
    <title></title>
    <link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
    <style type="text/css">
<!--
body {
	background-color: #eff8e9;
}
.style3 {color: #0d7d1e}
.style5 {font-size: small}
-->
    </style></head>
    <body>
	<!-- IF CALLBACK -->
	<%
		'INIT==================================================================
		IF Request.QueryString("ispostback") <> "" THEN
			Session.Contents("ispostback") = Request.querystring("ispostback")
		END IF 
		
		IF Session.Contents("ispostback") = "0" THEN
			Session.Contents("changeUnit") = Request.querystring("ictunit")
		END IF
		
		IF Request.QueryString("topCat") <> "" THEN
			Session.Contents("topCat") = Request.querystring("topCat")
		END IF
		
		'Response.Write(Cstr(Session.Contents("ispostback")))
		IF Request.querystring("icuitem") <> "" THEN
			Session.Contents("cuitem") = Request.querystring("icuitem")
		END IF
	
		IF Request.querystring("ictunit") <> "" THEN
			Session.Contents("ctunit") = Request.querystring("ictunit")			
		END IF 
	
		IF Request.querystring("iBaseDSD") <> "" THEN
			Session.Contents("BaseDSD") = Request.querystring("iBaseDSD")
		END IF
	
		IF Request.querystring("changeUnit") <> "" THEN
			Session.Contents("changeUnit") = Request.querystring("changeUnit")
		END IF
				
		IF Session.Contents("ispostback") = "0" THEN
			Session.Contents("ispublic") = "N"						
		END IF  
		IF Request.querystring("ispublic") <> "" THEN
			Session.Contents("ispublic") = Request.querystring("ispublic")	
		END IF
		
	'Response.Write(Session.Contents("ctunit")&",")
	'Response.Write(Session.Contents("changeUnit"))
		
		IF (Request.Form("submit")<>"") THEN
			IF (session("userid")<>"") THEN
				if(Session.Contents("topCat") = "") then
					Response.Write("<SCRIPT LANGUAGE='JAVASCRIPT'> alert('請選擇資料大類後再執行此功能')</SCRIPT>")
				else
					SQL="EXEC sp_move_document " & Cstr(Session.Contents("cuitem")) & "," + Cstr(Session.Contents("changeUnit")) &_
					",'" & Cstr(Session.Contents("ispublic")) & "','" & Cstr(Session.Contents("topCat")) & "' "
					conn.execute(SQL)
					SQL = "SELECT IDENT_CURRENT('CuDTGeneric') AS XX"				
					SET rs = conn.execute(SQL)
					xNewIdentity = rs("XX")
	
					'建立文件索引
					if checkGIPconfig("hyftdGIP") then
						hyftdGIPStr=hyftdGIP("add",xNewIdentity)
					end if	
					
					Response.Redirect("DsdXMLList.asp?iBaseDSD="& Request.Form("iBaseDSD"))
				end if
				'Response.Write("<SCRIPT LANGUAGE='JAVASCRIPT'> alert('" & Cstr(Session.Contents("cuitem")) & "," & Cstr(Session.Contents("changeUnit")) & "," & Cstr(Session.Contents("ispublic")) & "," & Cstr(Session.Contents("topCat")) & "')</SCRIPT>")
			END IF
		END IF
	%>
	
	<!-- =========== -->
<%
		'----1.1xmlSpec檔案檢查
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CtUnitX" & Cstr(Session.Contents("changeUnit")) & ".xml")
	if not fso.FileExists(LoadXML) then
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & Cstr(Session.Contents("BaseDSD")) & ".xml")
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
	
	xformClientCatRefLookup=nullText(oxml.selectSingleNode("//DataSchemaDef/dsTable/fieldList/field[fieldName='"+nullText(oxml.selectSingleNode("//DataSchemaDef/formClientCat"))+"']/refLookup"))
%>

	<!-- =========== -->
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><font size=2 color="#000000"><span class="style3">【<%=HTProgFunc%>】</span></td>
	  </tr>
	  <%if HTProgRight > 0 then %>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
		 <form id="form1" name="form1" method="post" action="ChangeUnit.asp?submit=1&ispostback=1">
		  <table width="95%" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bgcolor="#FFFFFF" class="TableBodyTD">
            <tr>
              <td bgcolor="#FFFFFF" class="TableHeadTD" font size=2 color="#000000"><div align="left" class="TableHeadTD">目前的主題單元</div></td>
              <td class="TableHeadTD" font size=2 color="#000000"> <div align="left" class="TableHeadTD">變更後的主題單元</div></td>
            </tr>
            <tr>
              <td align="left" valign="top">
			  	  <%
					SQL="SELECT cu.CtUnitID,cu.CtUnitName,cu.iBaseDSD FROM CtUnit cu WHERE cu.CtUnitID=" & Session.Contents("ctunit")
					SET RSS=conn.execute(SQL)
				  %>	
			  <% =RSS("CtUnitName")%></td>
              <td>主題單元:
                <%
					SQL="SELECT cu.CtUnitID,cu.CtUnitName,cu.iBaseDSD FROM CtUnit cu WHERE cu.CtUnitKind NOT IN ('U','1')"
					SET RSS=conn.execute(SQL)
				  %>
                  <select name="unitlist" id="unitlist" class="TableBodyTD" onChange="unitList_onchange()">
				  <% WHILE NOT RSS.EOF %>
                    <option value="<%=RSS("CtUnitID")%>" <% IF CStr(RSS("CtUnitID")) = CStr(Session.Contents("changeUnit")) THEN %> selected="selected" <% END IF %>>
					<%=RSS("CtUnitName")%>
					</option>
				  <% RSS.movenext 
				  WEND %>
				  
				  <% IF Session.Contents("changeUnit") <> "" THEN %>
                  </select>
                  <br>
                是否公開:
				  <select name="ispublic" class="TableBodyTD" onChange="ispublic_onchange">
				    <option value="Y" <% IF session("ispublic") = "Y" THEN %>selected<% END IF %>>公開</option>
                    <option value="N" <% IF session("ispublic") = "N" THEN %>selected<% END IF %>>不公開</option>
                  </select>
				  	<% IF xformClientCatRefLookup <>"" THEN %>
				   		<br>
		   		資料大類:
				<select name="datatype" class="TableBodyTD" onChange="datatype_onchange">
						<option value="" <% IF session("topCat") = "" THEN %>selected<% END IF %>>
						請選擇
						</option>
						<%
						SQL="SELECT codeMetaID, mCode, mValue, mSortValue FROM CodeMain WHERE codeMetaID = '"& xformClientCatRefLookup &"'"
						SET RSS=conn.execute(SQL)
						%>
						<% WHILE NOT RSS.EOF %>
						<option value="<%=RSS("mCode")%>" <% IF session("topCat") = RSS("mCode") THEN %>selected<% END IF %>>
						<%=RSS("mValue")%>
						</option>
						<% RSS.movenext 
						WEND %>
			   	</select>
					<% END IF %>
                    <br>
                    <input name="Submit" type="submit" class="TableBodyTD" value="儲存變更" onClick="save()"> 
				  <% END IF %>
			  </td>
            </tr>
			<%end if %>
          </table>
		  </form>
	    </label></td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
	       <a href="Javascript:window.history.back();">回前頁</a>	    
		</td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    
<%end sub '--- showHTMLHead() ------%>

<%sub ShowHTMLTail() %>      
    </td>
  </tr>  
</table> 
</body>
</html>
<%end sub '--- showHTMLTail() ------%>

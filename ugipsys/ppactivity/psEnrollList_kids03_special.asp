<% Response.Expires = 0
HTProgCap="�����W��"
HTProgFunc="�����ǭ��M��"
HTProgCode="PA005"
HTProgPrefix="psEnroll" %>
<% response.expires = 0 
Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open session("ODBCDSN")
Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------

%>

<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
	if request.queryString("paSID")<> "" then _
		session("paSID") = request.queryString("paSID")
 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 fSql=Request.QueryString("strSql")

if fSql="" then
	fSql = "SELECT htx.ppname, htx.pParentID, htx.ppbirth, htx.p2name, htx.psn2ID, htx.p2birth, htx.paSID, htx.ckValue, htx.erDate, htx.ppcount, htx.status, htx.psnID AS xPsnID, p.*" _
		& ", xref1.mValue AS xrStatus" _
		& " FROM (paEnroll AS htx LEFT JOIN paPsnInfo AS p ON p.psnID = htx.psnID)" _
		& " LEFT JOIN CodeMain AS xref1 ON xref1.mCode = htx.Status AND xref1.codeMetaID='ppCEstatus'" _
		& " WHERE htx.pasID=" & session("paSID")
	if request.form("htx_coName") <> "" then
		whereCondition = replace("htx.coName LIKE '%{0}%'", "{0}", request.form("htx_coName") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_coDesc") <> "" then
		whereCondition = replace("htx.coDesc LIKE '%{0}%'", "{0}", request.form("htx_coDesc") )
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_coDateS") <> "" then
		rangeS = request.form("htx_coDateS")
		rangeE = request.form("htx_coDateE")
		if rangeE = "" then	rangeE=rangeS
		whereCondition = replace("((htx.coBDate BETWEEN '{0}' AND '{1}') OR ('{0}' BETWEEN htx.coBDate AND htx.coEDate))", "{0}", rangeS)
		whereCondition = replace(whereCondition, "{1}", rangeE)
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_ceDateS") <> "" then
		rangeS = request.form("htx_ceDateS")
		rangeE = request.form("htx_ceDateE")
		if rangeE = "" then	rangeE=rangeS
		whereCondition = replace("((htx.ceBDate BETWEEN '{0}' AND '{1}') OR ('{0}' BETWEEN htx.ceBDate AND htx.ceEDate))", "{0}", rangeS)
		whereCondition = replace(whereCondition, "{1}", rangeE)
		fSql = fSql & " AND " & whereCondition
	end if
	if request.form("htx_status") <> "" then
		whereCondition = replace("htx.status = '{0}'", "{0}", request.form("htx_status") )
		fSql = fSql & " AND " & whereCondition
	end if
	for each x in request.form
	    if request(x) <> "" then
            if mid(x,2,3) = "fx_" then
	    	   select case left(x,1)
		          case "s"
		            	fSql = fSql & " AND " & mid(x,5) & "=" & pkStr(request(x),"")
		          case else
		            	fSql = fSql & " AND " & mid(x,5) & " LIKE '%" & request(x) & "%'"
		         end select
	        end if
	    end if
	next
	fSql = fSql & " ORDER BY htx.erdate"
end if

nowPage=Request.QueryString("nowPage")  '�{�b����

'response.write fSql & "<HR>"
'response.end

'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '�`����
   if totRec>0 then 
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=20  
      end if 
      
      RSreg.PageSize=PerPageSize       '�C������

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '�`����
      strSql=server.URLEncode(fSql)
   end if    
end if   
%>

<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Content-Language" content="zh-tw">
<title>��ߦ۵M��ǳժ��]</title>

<meta name="ProgId" content="FrontPage.Editor.Document">

<style type="text/css">
<!--
.txt       { font-family: "�s�ө���", "�ө���"; color:black; font-size: 11pt; line-height: 140%; letter-spacing: 1pt;}
.txt12       { font-family: "�s�ө���", "�ө���"; color:white; font-size: 12pt; line-height: 150%; letter-spacing: 1pt;}
.eng       { font-family: "Arial"; font-size: 90pt; line-height: 130%; letter-spacing: 1pt;; color: white}
.tt       { font-family: "�s�ө���", "�ө���"; font-size: 12pt; line-height: 120%; letter-spacing: 1pt;; color: white}
.short       { font-family: "�s�ө���", "�ө���"; font-size: 10pt; line-height: 60%; letter-spacing: 1pt;; color: white}
A          { font-family: "�s�ө���", "�ө���"; font-size: 10pt;color:white; line-height: 150%;}
A:hover    { text-decoration: underline; background-color: yellow;text-decoration: none; color:#336699}

-->
</style>

</head>



<body topmargin="0" leftmargin="0" bgcolor="#F5F3D3"  >
<table border="0" width="780" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" class=txt height="21" background="../uweb/scifamily/images/bg.gif"  align=center >
  <tr valign=top height=70 bgcolor=#336699>
     <td>
      <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="780" height="82" >
      <param name=movie value="../uweb/southbeach/images/title1.swf">
      <param name=quality value=high>
       <param name="wmode" value="transparent">
      <embed src="../uweb/southbeach/images/title1.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="780" height="70" wmode="transparent"> 
       </object>
     
     </td>
  </tr>
  <tr>
    <td width="780" valign="top" height="1" align="left">
       <p style="margin-top: -11; margin-bottom: -6"><img border="0" src="../scifamily/images/line.gif" width="800" height="8">
    </td>
  </tr>
  <tr >
    <td width="780" valign="top" height="12" bgcolor="#006699" align=right>
  
    </td>
  </tr>
  <tr>
    <td  valign="top" align="center"><br>&nbsp;
      <table border="0 width="90%" cellspacing="0" cellpadding="0" align=center height=1000 >
        <tr valign=top>
          <td width="100%" class=txt>
		<%
		sql = "SELECT htx.*, a.actName, a.actCat, a.actDesc, a.actTarget, xref2.mValue AS xrActCat" _
		& ", (SELECT count(*) FROM paEnroll AS e WHERE e.pasID=htx.pasID) AS erCount" _
		& ", (SELECT sum(ppCount) FROM paEnroll AS p WHERE p.pasID=htx.pasID) AS xppCount" _
		& " FROM paSession AS htx JOIN ppAct AS a ON a.actID=htx.actID" _
		& " LEFT JOIN CodeMain AS xref2 ON xref2.mCode = a.actCat AND xref2.codeMetaID='ppActCat'" _
		& " WHERE htx.paSID=" & session("paSID") 
		set RS = conn.execute(sql)
		%> �i<%=RS("actName")%>&nbsp;<%=RS("dtNote")%>�j

		
	<TABLE width=100% cellspacing="1" cellpadding="5" class=bg border=0 bgcolor="#9E946A" >                   
     		<tr align=center bgcolor="#F1EEE0" >    
		<td >�s��</td>
 		<td >�m�W</td>
		<td >�ǮթΪA�ȳ��</td>		
		<td class=lightbluetable>������</td>
		<td class=lightbluetable>�X�ͦ~���</td>
 		<td class=lightbluetable>�a���m�W</td>
 		<td class=lightbluetable>�a��������</td>
 		<td class=lightbluetable>�a���ͤ�</td> 		
 		<td class=lightbluetable>�s���q��</td> 		
 		<td class=lightbluetable>�P��ൣ�m�W</td>
		<td class=lightbluetable>�P��ൣ������</td>
		<td class=lightbluetable>�P��ൣ�ͤ�</td>
   		</tr>	
   		                                
		<%
		total = 0
    		for i=1 to 25                  
		%>                  
		<tr bgcolor="#F1EEE0" >                  
		<%pKey = ""
		pKey = pKey & "&psnID=" & RSreg("xPsnID")
		pKey = pKey & "&paSID=" & RSreg("paSID")
		
		if pKey<>"" then  pKey = mid(pKey,2)%>
		<TD bgcolor="#F1EEE0"  align=center>
		
		<% 
		total = total + Rsreg("ppcount")
		' response.write Rsreg("ppcount") 
		' response.write total
		' response.write RS("pbackup")
		if total < RS("pbackup") then 
		       
			 response.write i
		else
			 response.write "�ƨ�" 
                end if %>  			 
		
		</td>
	
		<TD  align=center nowrap>
		<font size=2><%=RSreg("pName")%>
		</font></td>
		<TD  align=center>
		<font size=2><%=RSreg("myOrg")%>
		</font></td>
	   <TD  align=center>
		<font size=2><%=RSreg("psnID")%></font></td>
       <TD  align=center>
		<font size=2><%=d7Date(RSreg("birthDay"))%>
		</font></td>
		<TD  align=center>
		<font size=2 ><%=RSreg("ppName")%>
		</font></td>
		<TD  align=center>
		<font size=2><%=RSreg("pParentID")%>
		</font></td>
		<TD  align=center>
		<font size=2><%=RSreg("ppbirth")%>
		</font></td>	
		<TD  align=center>
		<font size=2><%=RSreg("tel")%>
		</font></td>		
		<TD  align=center>
		<font size=2 ><%=RSreg("p2Name")%>
		</font></td>
		<TD  align=center>
		<font size=2 ><%=RSreg("psn2ID")%></font></td>
		<TD  align=center>
		<font size=2><%=RSreg("p2birth")%>
		</font></td>
		</tr>
    		<%
         	RSreg.moveNext
         	if RSreg.eof then exit for 
    		next 
		%>
    		</TABLE>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>





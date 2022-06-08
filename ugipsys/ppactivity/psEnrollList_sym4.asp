<% Response.Expires = 0
HTProgCap="錄取名單"
HTProgFunc="錄取學員清單"
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
	fSql = "SELECT htx.paSID, htx.ckValue, htx.erDate, htx.status, htx.psnID AS xPsnID, p.*" _
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

nowPage=Request.QueryString("nowPage")  '現在頁數

'response.write fSql & "<HR>"
'response.end

'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=20  
      end if 
      
      RSreg.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
      strSql=server.URLEncode(fSql)
   end if    
end if   
%>

<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Content-Language" content="zh-tw">
<title>國立自然科學博物館「演化相關概念在中小學自然科課程中的教學」座談會</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<style type="text/css">
<!--
.txt       { font-family: "新細明體", "細明體"; color:black; font-size: 11pt; line-height: 140%; letter-spacing: 1pt;}
.txt12       { font-family: "新細明體", "細明體"; color:white; font-size: 12pt; line-height: 150%; letter-spacing: 1pt;}
.eng       { font-family: "Arial"; font-size: 90pt; line-height: 130%; letter-spacing: 1pt;; color: white}
.tt       { font-family: "新細明體", "細明體"; font-size: 12pt; line-height: 120%; letter-spacing: 1pt;; color: white}
.short       { font-family: "新細明體", "細明體"; font-size: 10pt; line-height: 60%; letter-spacing: 1pt;; color: white}
A          { font-family: "新細明體", "細明體"; font-size: 10pt;color:white; line-height: 150%;}
A:hover    { text-decoration: underline; background-color: yellow;text-decoration: none; color:#336699}

-->
</style>

</head>



<body topmargin="0" leftmargin="0" bgcolor="#F5F3D3"  >
<table border="0" width="780" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" class=txt height="21" background="../uweb/symposium/images/bg.gif"  align=center >
  <tr valign=top height=70>
    <td width="780"  bgcolor="#0099CC" valign="top" height="82">
    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="780" height="82" >
      <param name=movie value="../uweb/symposium/images/title.swf">
      <param name=quality value=high>
       <param name="wmode" value="transparent">
      <embed src="../uweb/symposium/images/title.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="780" height="70" wmode="transparent"> 
       </object>
     </td>
  </tr>
  <tr>
    <td width="780" valign="top" height="1" align="left">
       <p style="margin-top: -11; margin-bottom: -6"><img border="0" src="..//symposium/images/line.gif" width="800" height="8">
    </td>
  </tr>
  <tr >
    <td width="780" valign="top" height="12" bgcolor="#006699" align=right>
         <a href="../uweb/symposium/index.htm">活動說明</a>｜<a href="../uweb/symposium/content.htm">座談內容</a>｜<a href="../uweb/symposium/method.htm">報名方式</a>｜<a href="../uweb/pact03sym4.asp">我要報名</a>｜<a href="http://info.nmns.edu.tw/WebNmns/item2-3.asp">活動訊息&nbsp;</a> 
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
		& " FROM paSession AS htx JOIN ppAct AS a ON a.actID=htx.actID" _
		& " LEFT JOIN CodeMain AS xref2 ON xref2.mCode = a.actCat AND xref2.codeMetaID='ppActCat'" _
		& " WHERE htx.paSID=" & session("paSID") 
		set RS = conn.execute(sql)
		%> 【<%=RS("actName")%>&nbsp;<%=RS("dtNote")%>】

		
	<TABLE width=100% cellspacing="1" cellpadding="5" class=bg border=0 bgcolor="#9E946A" >                   
     		<tr align=center bgcolor="#F1EEE0" >    
		<td >編號</td>
		
 		<td >姓名</td>
		<td >學校或服務單位</td>
		
   		</tr>	                                
		<%

    		for i=1 to 120                  
		%>                  
		<tr bgcolor="#F1EEE0" >                  
		<%pKey = ""
		pKey = pKey & "&psnID=" & RSreg("xPsnID")
		pKey = pKey & "&paSID=" & RSreg("paSID")
		
		if pKey<>"" then  pKey = mid(pKey,2)%>
		<TD bgcolor="#F1EEE0"  align=center>
		<% response.write i %>  			 
		</td>
	
		<TD  align=center nowrap>
		<font size=2><%=RSreg("pName")%>
		</font></td>
		<TD  align=center><font size=2 title="<%=RSreg("myOrg")%>">
		<%=RSreg("myOrg")%>
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





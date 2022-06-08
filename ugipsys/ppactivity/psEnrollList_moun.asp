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

<title>國立自然科學博物館-合歡山野外研習活動</title>


<style type="text/css">
<!--
.txt       { font-family: "新細明體", "細明體"; color:#336699; font-size: 10pt; line-height: 130%; letter-spacing: 1pt;}
.txt11       { font-family: "新細明體", "細明體"; color:#525858; font-size: 11pt; line-height: 150%; letter-spacing: 1pt;}
.eng       { font-family: "Arial"; font-size: 10pt; line-height: 130%; letter-spacing: 1pt;; color: #336699}
.tt       { font-family: "新細明體", "細明體"; font-size: 12pt; line-height: 120%; letter-spacing: 1pt;; color: white}
A          { font-family: "新細明體", "細明體"; font-size: 10pt;color:#336699; text-decoration: underline;line-height: 150%; }
A:hover    { text-decoration: underline;; background-color: #ffcc99;text-decoration: none;}

-->
</style>


</head>


<body bgcolor="#ffcc66"  topmargin="0" leftmargin="0">
<div align="center">
  <center>
  <table border="0" width="780" cellspacing="0" cellpadding="0" height="666" background="../uweb/mountain/images/bg.gif">
    <tr>
      <td width="249" height="18"></td>
      <td width="527" height="18"></td>
    </tr>
    <tr>
      <td width="776" height="63" colspan="2" bgcolor="#688C90" valign=top>
       <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="780" height="66" >
              <param name=movie value="../uweb/mountain/images/title.swf">
              <param name=quality value=high>
              <param name="wmode" value="transparent">
              <embed src="../uweb/mountain/images/title.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="780" height="66" wmode="transparent" >
              </embed> 
         </object> 
      </td>
    </tr>
    <tr>
      <td width="249" height="385" valign="top">
        <div align="center">
          <center>
          <table border="0" width="88%" height="362">
            <tr>
              <td width="100%" height="36">　</td>
            </tr>
            <tr>
              <td width="100%" height="131" valign="middle" align="center">
                <div align="center">
                  <center>
                <table border="1" width="196" height="122" bordercolor="#C0C0C0">
                  <tr >
                    <td valign="middle" align="center">
                      <applet code="fprotate.class" codebase="./" width="186" height="120">
                        <param name="rotatoreffect" value="dissolve">
                        <param name="time" value="5">
                        <param name="image1" valuetype="ref" value="../uWeb/mountain/images/02.jpg">
                        <param name="image2" valuetype="ref" value="../uWeb/mountain/images/03.jpg">
                        <param name="image3" valuetype="ref" value="../uWeb/mountain/images/04.jpg">
                        <param name="image4" valuetype="ref" value="../uWeb/mountain/images/05.jpg">
                      </applet>
                    </td>
                  </tr>
                </table>
                  </center>
                </div>
              </td>
            </tr>
            <tr>
              <td width="100%" height="183" valign="top">
                <div align="center">
                  <center>
                  <table border="0" width="80%" cellspacing="0" height="82">
                    <tr>
                      <td width="100%" height="36">
                        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="200" height="250" vspace="5" hspace="8" align="absmiddle">
				        <param name=movie value="../uweb/mountain/images/button.swf">
				        <param name=quality value=high>
				        <param name="wmode" value="transparent">
				        <embed src="../uweb/mountain/images/button.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="200" height="250" vspace="5" hspace="8" align="absmiddle" wmode="transparent">
				        </embed> 
			</object>
                      </td>
                    </tr>
                  </table>
                  </center>
                </div>
              </td>
            </tr>
          </table>
          </center>
        </div>
      </td>
      <td width="527" height="385" valign="top" align="center"><br>　
       <div align="center">
 	<table border="1" width="94%"  bordercolor="#C0C0C0" height="528">
 	        <tr valign=top>
          <td width="100%" class=txt>
		<%
		sql = "SELECT htx.*, a.actName, a.actCat, a.actDesc, a.actTarget, xref2.mValue AS xrActCat" _
		& ", (SELECT count(*) FROM paEnroll AS e WHERE e.pasID=htx.pasID) AS erCount" _
		& " FROM paSession AS htx JOIN ppAct AS a ON a.actID=htx.actID" _
		& " LEFT JOIN CodeMain AS xref2 ON xref2.mCode = a.actCat AND xref2.codeMetaID='ppActCat'" _
		& " WHERE htx.paSID=" & session("paSID") 
		set RS = conn.execute(sql)
		%> 　　　【<%=RS("actName")%>&nbsp;<%=RS("dtNote")%>】　<br>
        　
        <div align="center">
          <center>

		
	<TABLE width=407 cellspacing="1" cellpadding="5" class=bg border=0 bgcolor="#9E946A" >                   
     		<tr align=center bgcolor="#F1EEE0" >    
		<td width="65" >編號</td>
		
 		<td width="125" >姓名</td>
		<td width="173" >學校或服務單位</td>
		
   		</tr>	                                
		<%

    		for i=1 to 45                  
		%>                  
		<tr bgcolor="#F1EEE0" >                  
		<%pKey = ""
		pKey = pKey & "&psnID=" & RSreg("xPsnID")
		pKey = pKey & "&paSID=" & RSreg("paSID")
		
		if pKey<>"" then  pKey = mid(pKey,2)%>
		<TD bgcolor="#F1EEE0"  align=center width="65">
		<%if i< 41 then 
			 response.write i
		else
			 response.write "備取" 
                end if %>  			 
		</td>
	
	
	
	
	
	
		<TD  align=center nowrap width="125">
		<font size=2><%=RSreg("pName")%>
		</font></td>
		<TD  align=center width="173"><font size=2 title="<%=RSreg("myOrg")%>">
		<%=RSreg("myOrg")%>
		</font></td>
		</tr>
    		<%
         	RSreg.moveNext
         	if RSreg.eof then exit for 
    		next 
		%>

	</table>
          </center>
        </div>
      </td>
    </tr>
  </table>
  </center>
</div>
  </table>
</body>

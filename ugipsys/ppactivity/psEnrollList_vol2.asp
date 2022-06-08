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

<title>國立自然科學博物館劇場教室專案義工召募</title>
<style type="text/css">
<!--
.txt       { font-family: "新細明體", "細明體"; color:#336699; font-size: 10pt; line-height: 130%; letter-spacing: 1pt;}
.tt       { font-family: "新細明體", "細明體"; font-size: 11pt; line-height: 160%; letter-spacing: 1pt;; color: #336699}
A          { font-family: "新細明體", "細明體"; font-size: 10pt;color:#336699; text-decoration: underline;line-height: 150%; }
A:hover    { text-decoration: underline;; background-color: #ffcc99;text-decoration: none;}

-->
</style>
</head>

<BODY bgcolor="#7C849B" topmargin="0" leftmargin="0" text="#000099" >
  
<div align="center">
  <table border="0" width="780" cellspacing="0" cellpadding="0" >
    <tr>
      <td height="44" valign="top" bgcolor="#CCCCCC" align="center">
         <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="780" height="44" >
              <param name=movie value="../uweb/vol02/images/title.swf">
              <param name=quality value=high>
              <param name="wmode" value="transparent">
              <embed src="../uweb/vol02/images/title.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="780" height="44" wmode="transparent" >
              </embed> 
            </object> 
      </td>
    </tr>
  <center>
    <tr>
      <td height="14" valign="top" bgcolor="#D8DADE">
        <p style="line-height: 30%; margin-top: 0; margin-bottom: 0"><font size="1">　<br>
        </font></p>
        <p style="line-height: 30%; margin-top: 0; margin-bottom: 0">　</p>
      </td>
    </tr>
    <tr>
      <td height="174" valign="top" bgcolor="#D8DADE">
      <div align="center">
        <center>
        <table border="0" width="493" height="217" cellspacing="0" cellpadding="0">
          <tr>
            
            <td height="217" valign="top" width="239" align=center>
              <div align="center">
                <center>
                <table border="0" width="494" cellspacing="0" cellpadding="0" height="234" align=left>
                  <tr>
                    <td height="22" width="28"><img border="0" src="../uweb/dino/images/head_left.gif" width="28" height="30"></td>
                    <td width="429" height="22" bgcolor="#7C849B" align=center><font color=white>======== 
                      錄取名單 ========　</font></td>
                    <td width="31" height="22"><img border="0" src="../uweb/dino/images/head_right.gif" width="28" height="30"></td>
                  </tr>
                  <%
					sql = "SELECT htx.*, a.actName, a.actCat, a.actDesc, a.actTarget, xref2.mValue AS xrActCat" _
					& ", (SELECT count(*) FROM paEnroll AS e WHERE e.pasID=htx.pasID) AS erCount" _
					& " FROM paSession AS htx JOIN ppAct AS a ON a.actID=htx.actID" _
					& " LEFT JOIN CodeMain AS xref2 ON xref2.mCode = a.actCat AND xref2.codeMetaID='ppActCat'" _
					& " WHERE htx.paSID=" & session("paSID") 
					set RS = conn.execute(sql)
					%> 　
                  <p><font color="#800080"> 【<%=RS("actName")%>&nbsp;<%=RS("dtNote")%>】</font>
                  
                  
                  <tr>
                    <td width="485" valign="top" height="234" colspan="3" >
                      <table border="0" width="101%" bgcolor="#7C849B" height="318" cellspacing="1">
                        <tr >
                          <td width="100%" bgcolor="#D8DADE"  valign="top" align="left" class=txt height="318" class=txt>
                          　
                         
                         <TABLE width=359 cellspacing="1"  border=0 align=center>                   
  					   		<tr align=center>    
								<td width="59">編號</td>
						 		<td width="107">姓名</td>
								<td width="167" align="left">服務單位或學校</td>
							</tr>	
				    	
							<%  
								for i=1 to 60                  
							%>                  
							<tr class=tt>                  
								<%pKey = ""
								pKey = pKey & "&psnID=" & RSreg("xPsnID")
								pKey = pKey & "&paSID=" & RSreg("paSID")
								if pKey<>"" then  pKey = mid(pKey,2)%>
					    			<TD  align=center width="59">
								<%	 response.write i
									
						             %>  			 
								</td>
								<TD align=center nowrap width="107">
								 	<%=RSreg("pName")%>
								</td>
								<TD align=left width="167"><%=RSreg("myOrg")%>
								</td>
								
								
							</tr>
				    		<%
				         	RSreg.moveNext
				         	if RSreg.eof then exit for 
				    		next 
					   		%>
			    		</TABLE>

                         </td>
                       </tr>
                       <tr>
                        <td width="488" valign="top" height="1" colspan="3" >
                        <p style="line-height: 30%; margin-top: 0; margin-bottom: 0"><font size="1">　</font></p>
                       </td>
                      </tr>
                      <tr>
                       <td width="488" valign="top" height="1" colspan="3" bgcolor="#7C849B">
                        <p style="line-height: 30%; margin-top: 0; margin-bottom: 0"><font size="1">　</font></p>
                       </td>
                     </tr>
                    </table>
                  </td>
                 </tr>
            </table>
          </center>
        </div>
      </td>
    </tr>
    <tr>
      <td width="100%" valign="top" height="4" bgcolor="#D8DADE" class=txtsm>
        <p style="line-height: 40%; margin-top: 0; margin-bottom: 0"><font size="1">　</font></p>
      </td>
    </tr>
    <tr>
      <td width="100%" valign="top" height="211">
        <div align="center">
          <center>
          <table border="0" width="480" cellspacing="0" >
            <tr>
              <td width="56" height="46" valign="middle" align="center">
              <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="40" height="30" >
              <param name=movie value="../uweb/dino/images/white.swf">
              <param name=quality value=high>
              <param name="wmode" value="transparent">
              <embed src="../uweb/dino/images/white.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="40" height="30" wmode="transparent" >
              </embed> 
              </object> 

              </td>
              <td width="416" class=tt height="46" >國立自然科學博物館‧<span class="eng">National 
                Museum of Natural Science</span></td>
            </tr>
          </table>
          </center>
        </div>
      </td>
    </tr>
  </table>
  </center>
</div>
</body>

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
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5;no-caches;">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>�E�Q�G�~�ײĤ@�u�֦~��޳Ч@����笡��</title>
<style type="text/css">
<!--

.txt       { font-family: "�s�ө���", "�ө���"; font-size: 10pt; color: black; line-height: 175%; letter-spacing: 1pt;}
.tt       { font-family: "�s�ө���", "�ө���"; font-size: 9pt; color: black; line-height: 140%; letter-spacing: 2pt;}
.txt2       { font-family: "�s�ө���", "�ө���"; font-size: 11pt; color: black; line-height: 175%; letter-spacing: 1pt;}
.t1      { font-family: "�s�ө���", "�s�ө���"; font-size: 13pt; color: #660099; line-height: 180%; letter-spacing: 2pt; font-weight:bold;}
A          { color:#6600FF; text-decoration: underline; }
A:hover    { color:#6600FF; text-decoration: none; background-color: #F2DEFE;}

-->
</style>


</head>

<body bgcolor="#C4DCA5" text="#000000" background="../images/pc/bg.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="770" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
      <table width="99%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="106"><img src="../images/pc/cir.gif" width="106" height="94"></td>
          <td align="center"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="670" height="75">
              <param name=movie value="../images/pc//tit.swf">
              <param name=quality value=high>
              <param name="wmode" value="transparent">
              <embed src="../images/pc/tit.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="670" height="75" wmode="transparent">
              </embed> 
            </object></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td>
      <table width="99%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td class=txt align=center><font color="#378723">
           �����ʰt�X�Ǯաu�ͬ���ޡv���ҵ{�A���Ь�޳Ч@�������P�z��<br>�W�[�ꤤ�B��p�ǵ���Ĳ�ͬ���ެ��ʪ����|�A���տE�o����޳Ч@������C<br>
          </td>
        </tr>
        <tr>
          <td align="center"> 
            <table width="700" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="../images/pc/3.gif" width=100% height="10"></td>
              </tr>
              <tr>
                <td bgcolor="#EDF7E1">

	<table border="0" width="95%" cellspacing="1" cellpadding="0" align="center">
  	 <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;
		
	    </td>
  	</tr>
  	<tr>
    	    <td width="100%" colspan="2">
      		<hr noshade size="1" color="#000080">
	    </td>
  	</tr>
	<tr>
    		<td width="100%" colspan="2" class="FormRtext"></td>
  	</tr>
  	<tr>
    		<td width="95%" colspan="2" height=230 valign=top>
 		<Form name=reg method="POST" action=<%=HTprogPrefix%>List.asp>
  		<p align="center">  
		<!--<%If not RSreg.eof then%>     
     		<font size="2" color="rgb(63,142,186)"> ��
     		<font size="2" color="#FF0000"><%=nowPage%>/<%=totPage%></font>��|                      
        	�@<font size="2" color="#FF0000"><%=totRec%></font>
       		<font size="2" color="rgb(63,142,186)">��| ���ܲ�       
        	<select id=GoPage size="1" style="color:#FF0000">
             	<%For iPage=1 to totPage%> 
                   <option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
             	<%Next%>   
         	</select>      
         	��</font>           
       		<% if cint(nowPage) <>1 then %>             
            	|<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">�W�@��</a> 
       		<%end if%>      
       
        	<% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %> 
            	|<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">�U�@��</a> 
        	<%end if%>     
        	| �C������:
       	<select id=PerPage size="1" style="color:#FF0000">            
             <option value="10"<%if PerPageSize=10 then%> selected<%end if%>>10</option>                       
             <option value="20"<%if PerPageSize=20 then%> selected<%end if%>>20</option>
             <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
             <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
        </select>     
     </font>     
    <CENTER>-->
     	<TABLE width=95% cellspacing="1" cellpadding="5" class=bg border=0 bgcolor=336699 border=0>                   
     		<tr align=center>    
		<!--<td class=lightbluetable>�Ѧҭ�</td>-->
		<td class=lightbluetable>�s��</td>
		<td class=lightbluetable>������</td>
 		<td class=lightbluetable>�m�W</td>
		<td class=lightbluetable>�X�ͦ~���</td>
		<td class=lightbluetable>�NŪ�Ǯ�</td>
		<td class=lightbluetable>zipcode</td>
		<td class=lightbluetable>��}</td>
		<td class=lightbluetable>�q��</td>
		<td class=lightbluetable>Email</td>
		<td class=lightbluetable>����p���H</td>
   		</tr>	                                
		<%                   
    		for i=1 to 34                  
		%>                  
		<tr>                  
		<%pKey = ""
		pKey = pKey & "&psnID=" & RSreg("xPsnID")
		pKey = pKey & "&paSID=" & RSreg("paSID")
		if pKey<>"" then  pKey = mid(pKey,2)%>
		<TD class=whitetablebg align=center>
		<%if i<31 then 
			 response.write i
		else
			 response.write "�ƨ�" 
                end if %>  			 
		</td>
		<TD class=whitetablebg align=center>
		<font size=2><%=RSreg("psnID")%>
		</font></td>
		<TD class=whitetablebg align=center nowrap>
		<font size=2><%=RSreg("pName")%>
		</font></td>
		<TD class=whitetablebg align=center><font size=2 title="<%=RSreg("psnID")%>">
		<%=d7Date(RSreg("birthDay"))%>
		</font></td>
		<TD class=whitetablebg align=center><font size=2 title="<%=RSreg("myorg")%>">
		<%=RSreg("myorg")%>
		</font></td>
		<TD class=whitetablebg align=center><font size=2 title="<%=RSreg("zipcode")%>">
		<%=RSreg("zipcode")%>
		</font></td>
		<TD class=whitetablebg><font size=2 title="<%=RSreg("addr")%>">
		<%=RSreg("addr")%>
		</font></td>
		<TD class=whitetablebg><font size=2 title="<%=RSreg("tel")%>">
		<%=RSreg("tel")%>
		</font></td>
		<TD class=whitetablebg>
		<font size=2 title="<%=RSreg("eMail")%>">
		<%=RSreg("eMail")%>
		</font></td>
		<TD class=whitetablebg>
		<font size=2 title="<%=RSreg("emergContact")%>">
		<%=RSreg("emergContact")%>
		</font></td>
		</tr>
    		<%
         	RSreg.moveNext
         	if RSreg.eof then exit for 
    		next 
   		%>
    		</TABLE>
    		</CENTER>
       <!-- �{������ ---------------------------------->  
     </form>  
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center">
	</td></tr>
</table> 

        
        
        
        
                
                </td>
              </tr>
            </table>
          </td>
        </tr>
       </table>
     </td>
  </tr>         
</table>  
</form>
</body>
</html>                                 
<%else%>
      <script language=vbs>
           msgbox "�䤣����, �Э��]�d�߱���!"
'	       window.history.back
      </script>
<%end if%>

<script Language=VBScript>
  dim gpKey
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=" & newPage & "&strSql=" & strSql & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=<%=nowPage%>" & "&strSql=<%=strSql%>&pagesize=" & newPerPage                    
     end sub 

     sub setpKey(xv)
     	gpKey = xv
     end sub

	sub butAction(k) 
	  if gpKey = "" then 
	    alert("�Х���ܭӮ�!") 
	  else
    	select case k 
    	end select 
	  end if 
	end sub 

</script>

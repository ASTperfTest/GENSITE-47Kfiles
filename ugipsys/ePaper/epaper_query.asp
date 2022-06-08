<%@  codepage="65001" %>
<% Response.Expires = 0
HTProgCap="電子報管理"
HTProgFunc="電子報發行清單"
HTProgCode="GW1M51"
HTProgPrefix="ePub" 
%>
<!--#INCLUDE FILE="ePubListParam.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
    <meta name="GENERATOR" content="Hometown Code Generator 1.0">
    <title>查詢表單</title>
    <link type="text/css" rel="stylesheet" href="/css/list.css">
    <link type="text/css" rel="stylesheet" href="/css/layout.css">
    <link type="text/css" rel="stylesheet" href="/css/setstyle.css">
</head>
<body>
    <form name="iForm" method="post" action="epaper_resultList.asp?epubid=<%=request("epubid")%>&eptreeid=<%=Request("eptreeid")%>">
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr>
                <td width="100%" class="FormName" colspan="2">
                    電子報管理&nbsp;<font size="2">【最新農業知識文章查詢】</td>
            </tr>
            <tr>
                <td width="100%" colspan="2">
                    <hr noshade size="1" color="#000080">
                </td>
            </tr>
            <tr>
                <td class="FormLink" valign="top" colspan="2" align="right">
                    <a href="Javascript:window.history.back();">回前頁</a></td>
            </tr>
            <tr>
                <td class="Formtext" colspan="2" height="15">
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" width="80%" height="230" valign="top">
                    <center>
                        <table border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">
                            <tr>
                                <td class="eTableLable" align="right">
                                    資料庫來源：</td>
                                <td class="eTableContent">
                                    <select name="CtRootId" onchange="CtRootIdOnChange()">
                                        <option value="1" selected>入口網</option>
                                        <option value="2">主題館</option>
                                        <option value="3">知識家</option>
                                        <option value="4">知識庫</option>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td class="eTableLable" align="right">
                                    單元名稱：</td>
                                <td class="eTableContent">
                                    <input name="CtUnitName" size="30"></td>
                            </tr>
                            <tr>
                                <td class="eTableLable" align="right">
                                    資料標題：</td>
                                <td class="eTableContent">
                                    <label>
                                        <input name="sTitle" size="30"></label></td>
                            </tr>
                            <tr>
                                <td class="eTableLable" align="right">
                                    日期範圍：</td>
                                <td class="eTableContent">
                                    <input name="pcShowStartDate" size="8" readonly onclick="VBS: popCalendar 'StartDate',''">
                                    ～
                                    <input name="pcShowEndDate" size="8" readonly onclick="VBS: popCalendar 'EndDate',''">
                                    <input name="StartDate" type="hidden">
                                    <input name="EndDate" type="hidden">
                                </td>
                            </tr>
                            <tr>
                                <td class="eTableLable" align="right">
                                    狀態：</td>
                                <td class="eTableContent">
                                    <span class="FormLink">
                                        <select name="Status" size="1">
                                            <option value="Y">公開</option>
                                            <option value="N">不公開</option>
                                        </select>
                                    </span>
                                </td>
                            </tr>
                            <!--TR>
  				<TD class="eTableLable" align="right">單 位：</TD>
  				<TD class="eTableContent">
  					<Select name="DeptId" size=1>
    					<option value="" selected>不限定</option>
    					<%
    						'sql = "SELECT deptID, deptName FROM Dept WHERE (inUse = 'Y') ORDER BY deptID"
    						'Set Rs1 = Conn.Execute(sql)
    						'While Not Rs1.EOF
    						'	Response.Write "<option value=""" & Rs1("deptID") & """>" & Rs1("deptName") & "</option>"
    						'	Rs1.MoveNext
    						'Wend
    					%>
  					</select>
  				</TD>
				</TR-->
                        </table>
                    </center>
                    <table border="0" width="100%" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="100%">
                                <p align="center" />
                                <br />
                                <input name="submit" type="submit" value="查　詢" class="cbutton" />
                                <input name="reset" type="reset" value="重　填" class="cbutton" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <object data="../inc/calendar.asp" id="calendar1" type="text/x-scriptlet" width="245"
            height="160" style='position: absolute; top: 0; left: 230; visibility: hidden'>
        </object>
        <input type="hidden" name="CalendarTarget">
    </form>
</body>
</html>

<script language="javascript">
	<!--
	function CtRootIdOnChange()
	{		
		if(document.iForm.CtRootId.options[document.iForm.CtRootId.selectedIndex].value == 1) {
			document.getElementById("CtUnitName").disabled = false;
		}
		if(document.iForm.CtRootId.options[document.iForm.CtRootId.selectedIndex].value == 2) {
			document.getElementById("CtUnitName").disabled = false;
		}
		if(document.iForm.CtRootId.options[document.iForm.CtRootId.selectedIndex].value == 3) {			
			document.getElementById("CtUnitName").value = "";
			document.getElementById("CtUnitName").disabled = true;
		}
		if(document.iForm.CtRootId.options[document.iForm.CtRootId.selectedIndex].value == 4) {
			document.getElementById("CtUnitName").disabled = false;
		}
	}
	//-->
</script>

<script language="vbs">
Dim CanTarget
Dim followCanTarget

sub popCalendar(dateName,followName)        
 	CanTarget=dateName
 	followCanTarget=followName
	xdate = document.all(CanTarget).value
	if not isDate(xdate) then	xdate = date()
	document.all.calendar1.setDate xdate
	
 	If document.all.calendar1.style.visibility="" Then           
   		document.all.calendar1.style.visibility="hidden"        
 	Else        
       ex=window.event.clientX
       ey=document.body.scrolltop+window.event.clientY+10
       if ex>520 then ex=520
       document.all.calendar1.style.pixelleft=ex-80
       document.all.calendar1.style.pixeltop=ey
       document.all.calendar1.style.visibility=""
 	End If              
end sub     

sub calendar1_onscriptletevent(n,o) 
    	document.all("CalendarTarget").value=n         
    	document.all.calendar1.style.visibility="hidden"    
    if n <> "Cancle" then
        document.all(CanTarget).value=document.all.CalendarTarget.value
        document.all("pcShow"&CanTarget).value=document.all.CalendarTarget.value
        if followCanTarget<>"" then
        	document.all(followCanTarget).value=document.all.CalendarTarget.value
        	document.all("pcShow"&followCanTarget).value=document.all.CalendarTarget.value
        end if
	end if
end sub   
</script>


<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PublishQuery.aspx.cs" Inherits="PublishQuery" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>農漁生產作物管理</title>
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <link type="text/css" rel="stylesheet" href="../css/setstyle.css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <input name="clear" type="hidden" value="1" />
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr>
                <td width="50%" align="left" nowrap="nowrap" class="FormName">
                    農漁生產作物管理&nbsp;<font size="2"><asp:Literal ID="titles" runat="server"></asp:Literal></font>
                </td>
                <td width="50%" class="FormLink" align="right">
                    <a href="Publish.aspx?id=<%=pid%>" title="回前頁">回前頁</a>
                </td>
            </tr>
            <tr>
                <td width="100%" colspan="2">
                    <hr />
                    <br />
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" width="80%" height="230" valign="top">
                    <center>
                        <table width="100%" id="ListTable">
                            <tr>
                                <th>
                                    資料庫來源：
                                </th>
                                <td class="eTableContent">
                                    <asp:ListBox ID="CtRootId" runat="server" Rows="1">
                                        <asp:ListItem Value="1">入口網</asp:ListItem>
                                        <asp:ListItem Value="2">主題館</asp:ListItem>
                                        <asp:ListItem Value="3">知識家</asp:ListItem>
                                        <asp:ListItem Value="4">知識庫</asp:ListItem>
                                    </asp:ListBox>
                                </td>
                            </tr>
                            <tr>
                                <th>
                                    節點名稱：
                                </th>
                                <td class="eTableContent">
                                    <input name="CtNodeName" size="30" />(輸入節點名稱)
                                </td>
                            </tr>
                            <tr>
                                <th>
                                    資料標題：
                                </th>
                                <td class="eTableContent">
                                    <label>
                                        <input name="sTitle" size="30" />(輸入資料標題)</label>
                                </td>
                            </tr>
                            <tr>
                                <th>
                                    編修日期：
                                </th>
                                <td class="eTableContent">
								<object data="/inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
		<INPUT TYPE="hidden" name=CalendarTarget>
                                    <input name="xPostDateS" size="10" readonly onclick="VBS: popCalendar 'xPostDateS','xPostDateE'"> ～ 				
									<input name="xPostDateE" size="10" readonly onclick="VBS: popCalendar 'xPostDateE',''">
                                </td>
                            </tr>
                            <tr>
                                <th>
                                    狀態：
                                </th>
                                <td class="eTableContent">
                                    <span class="FormLink">
                                        <select name="Status" size="1">
                                            <option value="">請選擇</option>
                                            <option value="Y" selected="selected">公開</option>
                                            <option value="N">不公開</option>
                                        </select>
                                        <span style="color: #000000">(預設為公開)</span></span>
                                </td>
                            </tr>
                        </table>
                    </center>
                    <table border="0" width="100%" cellspacing="0" cellpadding="0">
                        <tr>
                            <td align="center">
                                <input type="submit" class="cbutton" value="查詢" />
                                <input type="reset" class="cbutton" value="重　填" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>
    </form>
<script language=vbs>
	cvbCRLF = vbCRLF
	cTabchar = chr(9)

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
			'document.all("x"&CanTarget).value=document.all.CalendarTarget.value
			if followCanTarget<>"" then
				document.all(followCanTarget).value=document.all.CalendarTarget.value
				'document.all("x"&followCanTarget).value=document.all.CalendarTarget.value
			end if
		end if
	end sub   
 	
	
</script> 	
</body>
</html>

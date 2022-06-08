<%@ Page Language="VB" Debug="true" %>
<%@ Import Namespace="System.Data" %>
<%@ Import NameSpace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="CuAttachAdd.asp">
<title>新增表單</title>
<script Language="VB" runat="Server">
Sub upload_File(sender as Object, e as EventArgs)
  Dim htx_path as String = request("htx_path")
  Dim htx_user as String = request("htx_aEditor")
  Dim htx_iCUItem as String = request("htx_xiCuItem")

  Dim up_path As String = Server.MapPath(htx_path)
 
  
  ' 取得HttpPostedFile物件
  Dim file As HttpPostedFile = htFile_NFileName.PostedFile
  ' 檢查檔案是否有內容
  If file.ContentLength = Nothing Then
    if request("htx_aTitle")=Nothing then
      msg.Text = "請輸入附件名稱及選擇上傳的檔案"
    else
      msg.Text = "請選擇上傳的檔案....."
    end if

  Else
    if request("htx_aTitle")=Nothing then
      msg.Text = "請輸入附件名稱"

    else
          dim ofname as string= Path.GetFileName(file.FileName)
	  dim fnext  as string= ""
	  if instrRev(ofname, ".")>0 then	
	    fnext=mid(ofname, instrRev(ofname, "."))
	  end if  
	  dim tstr as string= now()
	  dim nfname as string= right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  
    ' 上傳檔案
    file.SaveAs(up_path & "/" & nfname)
    
     Dim objCon As SqlConnection
     Dim objCmd As SqlCommand
     Dim strDbCon As String = "server=127.0.0.1;database=mGipmovie;uid=hyGIP;pwd=hyweb"
     Dim strSQL As String
     
     objCon = New SqlConnection(strDbCon)
     objCon.Open() ' 開啟資料庫連結
     objCmd = New SqlCommand(strSQL, objCon)
     objCmd.CommandText = "insert into CuDTAttach(xiCuItem,aTitle,aDesc,NFileName,aEditor,aEditDate,AttachkindA,Attachtype,bList,listSeq) values ('" & htx_iCUItem & "','" & request("htx_aTitle") & "','" & request("htx_aDesc") & "','" & nfname & "','" & htx_user & "',getdate(),'" & request("htx_AttachkindA") & "','" & request("htx_Attachtype") & "','" & request("htx_bList") & "','" & request("htx_listSeq") & "')"
     objCmd.ExecuteNonQuery()
     objCmd.CommandText = "insert into ImageFile(NewFileName,OldFileName) values ('" & nfname & "','" & ofname & "')"
     objCmd.ExecuteNonQuery()

     
     response.redirect("ftpsend.asp?action=add&iCUItem=" & htx_iCUItem & "&nfname=" & nfname)
     
    end if
  End If
End Sub
</script>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left">資料附件&nbsp;<font size=2>【新增】</td>
		<td width="50%" class="FormLink" align="right">
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A> 
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=90% height=230 valign=top>    

<form enctype="multipart/form-data" runat="Server">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER>
<TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">
<TR>
<TD class="eTableLable" align="right">
<font color="red">*</font>
附件名稱：</TD>
<TD class="eTableContent"><input id="htx_aTitle" size="50"  runat="Server"/>
<input type="hidden" name="htx_xiCuItem" value=<% =request("iCUItem") %>>
<input type="hidden" name="htx_aEditor" value=<% =request("user") %>>
<input type="hidden" name="htx_path" value=<% =request("path") %>>

</TD>
</TR>
<TR><TD class="eTableLable" align="right">說明：</TD>
<TD class="eTableContent"><input id="htx_aDesc" size="50"  runat="Server"/>
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
系統檔名：</TD>
<TD class="eTableContent"><input type="file" id="htFile_NFileName"  runat="Server"/>
</TD>
</TR>

<TR><TD class="eTableLable" align="right"><font color="red">*</font>附件檔案格式：</TD>
<TD class="eTableContent">
<Select id="htx_AttachkindA" runat="Server">
<option value="" selected>請選擇</option>
<option value="1">影音─WMV／AVI／MPG</option>
<option value="2">影音─MOV</option>
<option value="3">圖檔</option>
<option value="4">文字</option>
<option value="5">Flash檔</option>
</select>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">影音檔案類別：</TD>
<TD class="eTableContent">
<Select id="htx_Attachtype" runat="Server">
<option value="">請選擇</option>
<option value="1">寬頻</option>
<option value="2">窄頻</option>
</select>
</TD>
</TR>

<TR><TD class="eTableLable" align="right"><font color="red">*</font>
顯示於附件條列：</TD>
<TD class="eTableContent">
<select id="htx_bList" runat="Server">
<option value="Y" selected>是</option>
<option value="N">否</option>
</select>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">條列次序：</TD>
<TD class="eTableContent"><input id="htx_listSeq" size="2"  runat="Server"/>
</TD>
</TR>
</TABLE>
<asp:Label id="msg" Forecolor="red" runat="Server"/>
</CENTER>
    
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
               <asp:Button OnClick="upload_File" text="新增存檔" runat="Server"/>
               <input type="reset" name="reset" value="重填">
 </td></tr>
</table>
</form> 
    </td>
  </tr>  
</table> 
<SCRIPT LANGUAGE="vbs">
    sub setOtherCheckbox(xname,otherValue)
    	reg.all(xname).item(reg.all(xname).length-1).value = otherValue
    end sub
    sub doFileChoose()
    	window.open "CuAttachLarge.asp","","scrollbars=yes,width=560,height=400"
    end sub
    
    sub doAttachLarge(pickRadio)
    	reg.all("htFile_NFileName_Large").value = pickRadio
    end sub       
</SCRIPT>
</body>
</html>

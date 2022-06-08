<%@ Page Language="VB" Debug="true" %>
<%@ Import Namespace="System.Data" %>
<%@ Import NameSpace="System.Data.SqlClient" %>
<%@ Import Namespace="System.IO" %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="CuAttachEdit.asp">
<title>編修表單</title>
<script Language="VB" runat="Server">
Dim ConnString As String = ConfigurationSettings.appSettings("ConnString")

Sub update_File(sender as Object, e as EventArgs)
 Dim htx_path as String = request("htx_path")
  Dim htx_user as String = request("htx_aEditor")
  Dim htx_ixCuAttach as String = request("htx_ixCuAttach")
  Dim htx_xicuitem as String = request("htx_xicuitem")
  dim nfname as string
  
  Dim up_path As String = Server.MapPath(htx_path)
     Dim objCon As SqlConnection
     Dim objCmd As SqlCommand
     Dim strDbCon As String = ConnString
     Dim strSQL As String
     
     objCon = New SqlConnection(strDbCon)
     objCon.Open() ' 開啟資料庫連結
     objCmd = New SqlCommand(strSQL, objCon)
     
     objCmd.CommandText = "Select * From GipSites Where GipSiteID=N'" & request("siteid") & "'"
     objDataReader = objCmd.ExecuteReader()
     objDataReader.Read()
     Dim strDbCon2 As String =trim(objDataReader.Item("GipSiteDBConn2"))
     
     objCon = New SqlConnection(strDbCon2)
     objCon.Open() ' 開啟資料庫連結
     objCmd = New SqlCommand(strSQL, objCon)
  
  ' 取得HttpPostedFile物件
  Dim file As HttpPostedFile = htFile_NFileName.PostedFile
  ' 檢查檔案是否有內容

    if request("htx_aTitle")=Nothing then
      msg.Text = "請輸入附件名稱"
    
    end if
   if request("htx_aTitle")=Nothing then
      msg.Text = "請輸入附件名稱"
   else 
   If file.ContentLength <>0 Then
    
          dim ofname as string= Path.GetFileName(file.FileName)
	  dim fnext  as string= ""
	  if instrRev(ofname, ".")>0 then	
	    fnext=mid(ofname, instrRev(ofname, "."))
	  end if  
	  dim tstr as string= now()
	  nfname= right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  
    ' 上傳檔案
    file.SaveAs(up_path & "/" & nfname)
    
    Dim file2 As File 
    file2.Delete(up_path & "/" & request("htx_nfilename"))
    
     objCmd.CommandText = "insert into ImageFile(NewFileName,OldFileName) values ('" & nfname & "','" & ofname & "')"
     objCmd.ExecuteNonQuery()
     objCmd.CommandText = "delete ImageFile where NewFileName=N'" & request("htx_nfilename")  & "'"
     objCmd.ExecuteNonQuery()
     'response.write ("insert into ImageFile(NewFileName,OldFileName) values ('" & nfname & "','" & ofname & "')")
    else
          nfname=request("htx_nfilename") 
    end if
    
     
     objCmd.CommandText = "update CuDTAttach set aTitle=N'" & request("htx_aTitle") & "',aDesc=N'" & request("htx_aDesc")  & "',NFileName=N'" & nfname & "',aEditor=N'" & htx_user & "',aEditDate=getdate(),bList=N'" & request("htx_bList") & "',listSeq=N'" & request("htx_listSeq") & "' where ixCuAttach=N'" & htx_ixCuAttach &"'" 
     objCmd.ExecuteNonQuery()
     
     response.redirect("ftpsend.asp?action=mod&iCUItem=" & htx_xicuitem & "&nfname=" & nfname & "&ofname=" & request("htx_nfilename"))
    
    end if
End Sub
Sub delete_File(sender as Object, e as EventArgs)
  Dim htx_path as String = request("htx_path")
  Dim htx_user as String = request("htx_aEditor")
  Dim htx_ixCuAttach as String = request("htx_ixCuAttach")
  Dim htx_xicuitem as String = request("htx_xicuitem")
  dim nfname as string
  
  Dim up_path As String = Server.MapPath(htx_path)
     Dim objCon As SqlConnection
     Dim objCmd As SqlCommand
     Dim strDbCon As String = ConnString
     Dim strSQL As String
     
     objCon = New SqlConnection(strDbCon)
     objCon.Open() ' 開啟資料庫連結
     objCmd = New SqlCommand(strSQL, objCon)
     
     objCmd.CommandText = "Select * From GipSites Where GipSiteID=N'" & request("siteid") & "'"
     objDataReader = objCmd.ExecuteReader()
     objDataReader.Read()
     Dim strDbCon2 As String =trim(objDataReader.Item("GipSiteDBConn2"))
     
     objCon = New SqlConnection(strDbCon2)
     objCon.Open() ' 開啟資料庫連結
     objCmd = New SqlCommand(strSQL, objCon)
  
     Dim file2 As File 
     file2.Delete(up_path & "/" & request("htx_nfilename"))
    
     objCmd.CommandText = "delete ImageFile where NewFileName=N'" & request("htx_nfilename")  & "'"
     objCmd.ExecuteNonQuery()
    
     objCmd.CommandText = "delete CuDTAttach where ixCuAttach=N'" & htx_ixCuAttach &"'" 
     objCmd.ExecuteNonQuery()
     
    response.redirect("ftpsend.asp?action=del&iCUItem=" & htx_xicuitem & "&nfname=" & request("htx_nfilename"))
    
End Sub

      dim xicuitem as string
      dim atitle as string
      dim adesc as string
      dim nfilename as string
      dim aeditor as string
      dim aeditdate as string
      dim blist as string
      dim listseq as string
      dim fxr_nfilename as string
      
     Dim objCon As SqlConnection
     Dim objCmd As SqlCommand
     Dim strDbCon As String = ConnString
     Dim objDataReader As SqlDataReader
     Dim strSQL As String

</script>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>

    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left">附件管理&nbsp;<font size=2>【編修】</td>
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
	    <td align=center colspan=2 width=80% height=230 valign=top>    
	    
<%

     
     objCon = New SqlConnection(strDbCon)
     objCon.Open() ' 開啟資料庫連結
      objCmd = New SqlCommand(strSQL, objCon)
     
     objCmd.CommandText = "Select * From GipSites Where GipSiteID=N'" & request("siteid") & "'"
     objDataReader = objCmd.ExecuteReader()
     objDataReader.Read()
     Dim strDbCon2 As String =trim(objDataReader.Item("GipSiteDBConn2"))
     
     objCon = New SqlConnection(strDbCon2)
     objCon.Open() ' 開啟資料庫連結
     objCmd = New SqlCommand(strSQL, objCon)
     objCmd.CommandText = "SELECT htx.xicuitem,htx.atitle,isnull(htx.adesc,'') adesc,htx.nfilename,htx.aeditor,htx.aeditdate,htx.blist,isnull(htx.listseq,'') listseq, isnull(xrefNFileName.oldFileName,'')  AS fxr_NFileName FROM (CuDTAttach AS htx LEFT JOIN imageFile AS xrefNFileName ON xrefNFileName.newFileName = htx.NFileName) WHERE htx.ixCuAttach=" & request("ixCuAttach")
     objDataReader = objCmd.ExecuteReader()
 
    while objDataReader.Read() 
      xicuitem=trim(objDataReader.Item("xicuitem"))
      atitle=trim(objDataReader.Item("atitle"))
      adesc=trim(objDataReader.Item("adesc"))
      nfilename=trim(objDataReader.Item("nfilename"))
      aeditor=trim(objDataReader.Item("aeditor"))
      aeditdate=trim(objDataReader.Item("aeditdate"))
      blist=trim(objDataReader.Item("blist"))
      listseq=trim(objDataReader.Item("listseq"))
      fxr_nfilename=trim(objDataReader.Item("fxr_nfilename"))
     
       end while
     objDataReader.Close()
%>   
<form enctype="multipart/form-data" runat="Server"> 
<INPUT TYPE=hidden name=submitTask value="">
<input type="hidden" name="htx_xicuitem" value="<% =xicuitem %>">
<input type="hidden" name="htx_ixCuAttach" value="<% =request("ixCuAttach") %>">
<input type="hidden" name="htx_aEditor" value="<% =request("user") %>">
<input type="hidden" name="htx_path" value="<% =request("path") %>">
<input type="hidden" name="htx_nfilename" value="<% =nfilename %>">
<input type="hidden" name="htx_ofilename" value="<% =fxr_nfilename %>">

<CENTER>
<TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">
<TR>
<TD class="eTableLable" align="right"><font color="red">*</font>
附件名稱：</TD>
<TD class="eTableContent"><input name="htx_aTitle" value="<% =atitle %>" >
</TD>
</TR>
<TR><TD class="eTableLable" align="right">說明：</TD>
<TD class="eTableContent"><input name="htx_aDesc" size="50" value="<% =adesc %>" >
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
系統檔名：</TD>
<TD class="eTableContent"><table border="0" cellpadding="0" cellspacing="0" width="100%">
	
	<tr>
	<td>
	<input type="file" id="htFile_NFileName"  runat="Server"/>
	<span id="logo_NFileName" class="rdonly"></span>
		<div id="noLogo_NFileName" style="color:red"><% =fxr_nfilename %></div></td>
	<td valign="bottom">
		
	</td></tr></table>
</TD>
</TR>
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
顯示於附件條列：</TD>
<TD class="eTableContent">
<select name="htx_bList">
<% if blist="Y" then  %>
<option value="Y" selected>是</option>
<option value="N" >否</option>
<% else %>
<option value="Y">是</option>
<option value="N"  selected>否</option>
<% end if %>
</select>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">條列次序：</TD>
<TD class="eTableContent"><input name="htx_listSeq" size="2" value="<% =listseq %>">
</TD>
</TR>
</TABLE>
<asp:Label id="msg" Forecolor="red" runat="Server"/>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
               <asp:Button OnClick="update_File" text="編修存檔" class="cbutton"  runat="Server"/>
               <asp:Button OnClick="delete_File" text="刪　除"  class="cbutton"  runat="Server"/>
                  
        <input type=reset value ="重　填" class="cbutton" >
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">
 </td></tr>
</table>
</form>   
    </td>
  </tr>  
</table> 
</body>
</html>

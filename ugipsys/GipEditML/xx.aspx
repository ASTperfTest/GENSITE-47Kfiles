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
Sub update_File(sender as Object, e as EventArgs)
	dim x as object
 response.write ("1." & request.form("xx"))
 for each x in request.form
 	response.write( "<BR/>" & x & ")->" & request.form(x))
 next
 'response.end
End Sub
Sub delete_File(sender as Object, e as EventArgs)
 
End Sub


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
	    <form enctype="multipart/form-data" runat="Server"> 
            <INPUT TYPE=hidden name=submitTask value="">     
<%
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
     Dim strDbCon As String = "server=10.10.5.59;database=mGIPmoj;uid=hyGIP;pwd=hyweb"
     Dim objDataReader As SqlDataReader
     Dim strSQL As String

     
     objCon = New SqlConnection(strDbCon)
     objCon.Open() ' 開啟資料庫連結
     objCmd = New SqlCommand(strSQL, objCon)
     objCmd.CommandText = "SELECT htx.xicuitem,htx.atitle,htx.adesc,htx.nfilename,htx.aeditor,htx.aeditdate,htx.blist,isnull(htx.listseq,'') listseq, xrefNFileName.oldFileName AS fxr_NFileName FROM (CuDTAttach AS htx LEFT JOIN imageFile AS xrefNFileName ON xrefNFileName.newFileName = htx.NFileName) WHERE htx.ixCuAttach=" & request("ixCuAttach")
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
     
%>
<INPUT name="xx" value="<%=atitle%>">
<%

    end while
     objDataReader.Close()

%> 

<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
               <asp:Button OnClick="update_File" text="編修存檔" runat="Server"/>
               <asp:Button OnClick="delete_File" text="刪　除" runat="Server"/>
                  
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

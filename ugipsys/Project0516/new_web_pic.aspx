<%@ Page Language="C#" AutoEventWireup="true" CodeFile="new_web_pic.aspx.cs" Inherits="new_web_pic" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>自訂標題圖片上傳</title>
    <link href="./css/list.css" rel="stylesheet" type="text/css">
<link href="./css/layout.css" rel="stylesheet" type="text/css">
<link href="./css/theme.css" rel="stylesheet" type="text/css">
    <script language="javascript" type="text/javascript">
   
    function check_login()
{
      
     var Title = document.getElementById("Txt_topic");
	 var file = document.form1.Banner_Upload.value;
     
	if(Title.value == "")
	{	
		alert("請輸入主題");
		event.returnValue = false;
		
	}
	else if(file == "")
	{
	     alert("請選擇圖片上傳");
		 event.returnValue = false;   	    
	}
	
	
}
    
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <table class="popup">
<caption>【自訂標題圖片上傳】</caption>
<tr>
<th>圖片標題：</th>
<td><asp:TextBox ID="Txt_topic" runat="server" CssClass="box" Width="209px"></asp:TextBox> </td>
</tr>
<tr>
<th>上傳設計圖片：</th>
<td>
    <asp:FileUpload ID="Banner_Upload" runat="server" Width="266px" CssClass="box" /><br>(建議圖片大小：px）</td>
</tr>

</table>

<div class="settingbutton">
<asp:Button ID="go" runat="server" Text="確定" OnClick="go_Click" OnClientClick="check_login()" />
<input name="Close2" type="button" class="tx1" onClick="self.close();return false" value="取消">
</div>
    
    </div>
    </form>
</body>
</html>

<%@ Page Language="C#" AutoEventWireup="true" CodeFile="editlogo.aspx.cs" Inherits="editlogo" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title><%=title%></title>
    <link rel="stylesheet" type="text/css" href="./css/css.css">
    <script type="text/javascript">
	     function opennew(url) {
             window.open(url, '', 'width=700,height=400,resizable=yes,toolbar=no');
         }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div id="matchinfo" runat="server">
        <div id="matchlogoA" runat="server">
            <h3>參賽作品</h3>
	    <a href='javascript:opennew("/logoselection/fullimage.aspx?fileName=<%=SubjectId%>.jpg")'>
			<asp:Image ID="ImageA" runat="server" Visible="false" width="200px" height="100px"/>
	    </a>	
			<br/>
            <asp:FileUpload ID="FileUploadA" runat="server" />
            <asp:Button ID="UploadlogoA" runat="server" Text="上傳" OnClick="UploadLogoA_Click"/>
            <br/>
        </div>
        <asp:Label ID="FileExt" runat="server" Text="注意:上傳檔案必須為副檔名jpg之RGB格式圖檔<br>若看不到上傳的圖檔,請檢查圖檔是否為CMYK格式,並改以RGB格式重新上傳<br>若重新上傳後仍然看到舊圖,請按CTRL+F5重新整理" ForeColor ="Red" Font-Size="X-Small"></asp:Label>
		<br/>
    </div>
    </form>
</body>
</html>

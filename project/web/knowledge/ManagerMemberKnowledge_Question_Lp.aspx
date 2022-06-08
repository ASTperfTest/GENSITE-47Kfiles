<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ManagerMemberKnowledge_Question_Lp.aspx.vb" Inherits="ManagerMemberKnowledge_Question_Lp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="include/style.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
 <table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
    <td class="content">
      <div class="content_header"><img src="image/header.gif" width="500" height="135" /></div>
      <div class="content_mid">
    <h3>活動發問列表</h3>
 <div >
        <asp:Label ID="LabelBackLink" runat="server" Text="" />  
			 
	 </div>
            <!--分頁 -->
		    <div class="Page">
        	    第<asp:Label ID="PageNumberText" runat="server" Text="" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server" Text="" CssClass="Number" />頁，
        	    共<asp:Label ID="TotalRecordText" runat="server" Text="" CssClass="Number" />筆資料，     
        	    <asp:HyperLink ID="PreviousLink" runat="server">
        	        <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg" Visible="false" AlternateText="上一頁"></asp:image>        	            
        	        <asp:Label ID="PreviousText" runat="server" Visible="false" Text="Label">上一頁 &nbsp;</asp:Label>            
        	    </asp:HyperLink>
                到第 <asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" /> 頁 &nbsp;
                <asp:HyperLink ID="NextLink" runat="server">
                    <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg" Visible="false" AlternateText="下一頁"></asp:image>                     	
                    <asp:Label ID="NextText" runat="server" Visible="false" Text="Label">下一頁 &nbsp;</asp:Label>
                </asp:HyperLink>，每頁顯示                      
                <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
                    <asp:ListItem Selected="True" Value="10">10</asp:ListItem>
                    <asp:ListItem Value="20">20</asp:ListItem>
                    <asp:ListItem Value="30">30</asp:ListItem>
                    <asp:ListItem Value="50">50</asp:ListItem>
                </asp:DropDownList> 筆
            </div>
            
            <!--表格條列 /開始 -->
            <asp:Label ID="TableText" runat="server" Text="" />  
<P/>			
    </div>
</td>
  </tr>
</table>
    </form>
</body>
</html>

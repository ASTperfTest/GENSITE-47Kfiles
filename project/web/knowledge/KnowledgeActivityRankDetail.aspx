<%@ Page Language="VB" AutoEventWireup="false" CodeFile="KnowledgeActivityRankDetail.aspx.vb" Inherits="KnowledgeActivityRankDetail" %>

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
     <%-- <asp:Label ID="TabText" runat="server" Text=""></asp:Label>--%>
  </tr>
  <tr>
    <td class="content">
      <div class="content_header"><img src="image/header.gif" width="500" height="135" /></div>
      <div class="content_mid">

			<P>
			<div class="Event">	
			<div >
			
		   篩選條件：<br/>
		   會員:<asp:TextBox ID="TextBoxMember" runat="server" Width="100px"></asp:TextBox><br/>
      得分:<asp:TextBox ID="TextBoxScoreS" runat="server" Width="30px"></asp:TextBox>~<asp:TextBox ID="TextBoxScoreE"
runat="server" Width="30px"></asp:TextBox>
<asp:RegularExpressionValidator id="validatorTextBoxScoreS" ValidationExpression="\d*" runat="server" ControlToValidate="TextBoxScoreS" ErrorMessage="得分(起)只能輸入數字" Display="None" SetFocusOnError="true" />
<asp:RegularExpressionValidator id="validatorTextBoxScoreE" ValidationExpression="\d*" runat="server" ControlToValidate="TextBoxScoreE" ErrorMessage="得分(迄)只能輸入數字" Display="None" SetFocusOnError="true"/>
<asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="true" ShowSummary="false"/>
<br/>
		  
<asp:Button ID="btnQuery" runat="server" Text="查詢" />
<asp:HyperLink ID="linkExport"  runat="server" Target="_blank">匯出報表</asp:HyperLink>
		   </div>
			  <div class="Page">
				第<asp:Label ID="PageNumberText" runat="server" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server"  CssClass="Number"/>頁，
				共<asp:Label ID="TotalRecordText" runat="server" CssClass="Number" />筆資料，     
				<asp:HyperLink ID="PreviousLink" runat="server">
				  <asp:image runat="server" ImageUrl="/xslGip/style1/images3/arrow_left.gif" ID="PreviousImg" AlternateText="上一頁" />
				  <asp:Label ID="PreviousText" runat="server" >上一頁 &nbsp;</asp:Label>
				</asp:HyperLink>
				跳至第<asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" /> 頁 &nbsp;
				<asp:HyperLink ID="NextLink" runat="server">
				  <asp:image runat="server" ImageUrl="/xslGip/style1/images3/arrow_right.gif" ID="NextImg" AlternateText="下一頁"/>
				  <asp:Label ID="NextText" runat="server">下一頁 &nbsp;</asp:Label>
				</asp:HyperLink>，每頁                      
				<asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
				  <asp:ListItem Selected="True" Value="15">15</asp:ListItem>                    
				  <asp:ListItem Value="30">30</asp:ListItem>
				  <asp:ListItem Value="50">50</asp:ListItem>
				</asp:DropDownList>筆資料
			  </div>

			<asp:Label ID="TableText" runat="server" Text="" />   	
		    </div>
			</p>
		</div></td>
  </tr>
</table>
    </div>
    </form>
</body>
</html>

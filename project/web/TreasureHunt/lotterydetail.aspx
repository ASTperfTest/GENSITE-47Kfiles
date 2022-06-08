<%@ Page Language="C#" AutoEventWireup="true" CodeFile="lotterydetail.aspx.cs" Inherits="TreasureHunt_lotterydetail" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>未命名頁面</title>
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <style type="text/css"> 
    .present_title{
	color:#F30;
	font-weight:bold;
	font-size:13px;
	white-space:nowrap;
    }
    .chance{
	font-weight:bold;
	color:#F60;
	font-size:15px;
	text-decoration:underline;
	margin:0 3px;
    }
  </style> 
</head>
<body>
    <form id="form1" runat="server">
    <table width="100%"><tr><td class ="content" width="100%"><br/>
    <div class="pantology2">
    <div class="Page">
    <input type="button" value="回排行榜" onclick="location.href='/treasureHunt/activityrankdetail.aspx?avtivityid=<%=activityId %>'" /> 
	<br/>
	<h4>各抽獎獎項總兌換數量:</h4>
    <asp:label ID="LabelGiftVotes" runat="server" text=""></asp:label><br/>
     
     
     <asp:label ID="LabelTableName" runat="server" text=""></asp:label><br/><br/>
     
     篩選條件：<br/>
		   
		   <table><tr>
		                <td>會員:
		                </td>
		                <td><asp:TextBox ID="TextBoxMember" runat="server" Width="100px"></asp:TextBox>&nbsp;(帳號、姓名、暱稱)
		                </td>
		          </tr>
		          <tr>
		                <td>禮物:
		                </td>
		                <td><asp:DropDownList ID="GiftCategory" AutoPostBack="false" runat="server" /> 
		                </td>
		          </tr>
		           <tr>
		                <td colspan="2"><input type="submit" id="btnQuery" value="查詢" />&nbsp;&nbsp;<input type="button" id="clearAll" value="回首頁" onclick="window.location='/treasurehunt/lotterydetail.aspx?avtivityid=<%=activityId %>'"/>
		                </td>
		          </tr>
		  </table>
     
     
     <asp:HyperLink ID="linkExport"  runat="server" Target="_blank">匯出報表</asp:HyperLink><br />
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
        <asp:Label ID="TableText" runat="server" Text="" />  
    </div>
    </div>
    </td></tr></table>
    </form>
</body>
</html>

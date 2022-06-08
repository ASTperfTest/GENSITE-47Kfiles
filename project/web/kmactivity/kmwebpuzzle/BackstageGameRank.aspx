<%@ Page Language="C#" AutoEventWireup="true" CodeFile="BackstageGameRank.aspx.cs" Inherits="BackstageGameRank" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="/js/jquery-1.3.2.min.js"></script>

    <script type="text/javascript" src="/js/datepicker.js"></script>

    <link rel="Stylesheet" href="/css/datepicker.css" />
    <link type="text/css" href="/css/jquery.css" rel="stylesheet" />
    <script type="text/javascript" language="JavaScript">
        $(document).ready(function (e) {
            $("#TxtStartDate").datepicker({
                dateFormat: "yy/mm/dd"
            });
            $("#TxtEndDate").datepicker({
                dateFormat: "yy/mm/dd"
            });
        })
     </script>
</head>
<body>
<!--版面樣式大致上已經OK了-->
    <form id="form1" runat="server">
    <div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
             <td class="content">
             <div class="content_mid">
        <div class="Page">
			<br/>
		   篩選條件：<br/>
		   
		   <table><tr>
		                <td>會員:
		                </td>
		                <td><asp:TextBox ID="TextBoxMember" runat="server" Width="100px"></asp:TextBox>&nbsp;(帳號、姓名、暱稱)
		                </td>
		          </tr>
                  
		          <tr>
		                <td>分數:
		                </td>
		                <td><asp:TextBox ID="TextScoreLowerBound" runat="server" Width="30px"></asp:TextBox>~<asp:TextBox ID="TextScoreUpperBound" runat="server" Width="30px"></asp:TextBox>
		                <asp:RegularExpressionValidator id="RegularExpressionValidator1" ValidationExpression="\d*" runat="server" ControlToValidate="TextScoreUpperBound" ErrorMessage="獲得寶物數只能輸入數字" Display="None" SetFocusOnError="true"/>
		                <asp:RegularExpressionValidator id="RegularExpressionValidator2" ValidationExpression="\d*" runat="server" ControlToValidate="TextScoreLowerBound" ErrorMessage="獲得寶物數只能輸入數字" Display="None" SetFocusOnError="true"/>
		                </td>
		          </tr>
		          <tr>
		                <td>完成數:
		                </td>
		                <td>
                        <asp:TextBox ID="TextBox1" runat="server" Width="30px"></asp:TextBox>~<asp:TextBox ID="TextBox2" runat="server" Width="30px"></asp:TextBox>
		                </td>
		          </tr>
		          <tr>
		                <td>日期區間:
		                </td>
		                <td>
		                <asp:TextBox ID="TxtStartDate" runat="server" Width="80px" ></asp:TextBox>~<asp:TextBox ID="TxtEndDate" runat="server" Width="80px"></asp:TextBox>
		                
		                </td>
		          </tr>

		          <tr>
		                <td colspan="2">
                            <asp:Button ID="Button1" runat="server" Text="查詢" onclick="Button1_Click" />&nbsp;&nbsp;
		                </td>
		          </tr>
		   </table>
            <asp:HyperLink ID="linkExport"  runat="server" Target="_blank">匯出報表</asp:HyperLink> &nbsp;&nbsp;&nbsp;  
            <asp:Button ID="Button2" runat="server" onclick="Button2_Click" Text="改變使用者狀態" /> &nbsp;&nbsp;&nbsp;  
            <asp:LinkButton ID="linkCheckCheat" runat="server" OnClick="CheckCheat">驗證錯誤清單</asp:LinkButton>
		   </div>
           
        第<asp:Label ID="PageNumberText" runat="server" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server"  CssClass="Number"/>頁，
   	    共<asp:Label ID="TotalRecordText" runat="server" CssClass="Number" />筆資料，     
        <asp:LinkButton ID="preLink" runat="server" Visible="false" OnClick="preLinkAct">
            <asp:Label ID="preText" runat="server" Text="上一頁" />&nbsp;
        </asp:LinkButton>
        跳至第<asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" OnSelectedIndexChanged="ChangePageNumber" runat="server" /> 頁 &nbsp;
        <asp:LinkButton ID="nextLink" runat="server" Visible="false" OnClick="nextLinkAct">
            <asp:Label ID="nextText" runat="server" Text="下一頁" />
        </asp:LinkButton>
        ，每頁                      
        <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" OnSelectedIndexChanged="ChangePageSize" runat="server">
        <asp:ListItem Selected="True" Value="15">15</asp:ListItem>                    
        <asp:ListItem Value="30">30</asp:ListItem>
        <asp:ListItem Value="50">50</asp:ListItem>
        </asp:DropDownList>筆資料

        <asp:Repeater ID="rpList" runat="server">
            <HeaderTemplate>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <th>
                         &nbsp;
                        </th>
                        <th>
                            停權
                        </th>
                        <th>
                            帳號
                        </th>
                        <th>
                            姓名│暱稱
                        </th>
                        <th>
                            活動得分
                        </th>
                        <th>
                            完成數
                        </th>
                        <th>
                            使用點數
                        </th>
                        <th>
                            未使用點數
                        </th>
                        <th>
                            總點數
                        </th>
                    </tr>
            </HeaderTemplate>
            <ItemTemplate>
                <tr>
                    <td align="center">
                        <%#Eval("Row") %>
                    </td>   
                    <td align="center">
                        <asp:CheckBox ID="CheckBox1" runat="server"  Checked='<%#Convert.ToBoolean(Eval("status")) %>' /></asp:CheckBox>
                    </td>
                    <td align="center">
                        <asp:Label runat="server" ID="labmemberid" Text='<%#Eval("LOGIN_ID")%>' Visible="false"></asp:Label>
                        <a href="/kmactivity/kmwebpuzzle/UserGameLog.aspx?loginid=<%#Eval("LOGIN_ID")%>"><%#Eval("LOGIN_ID")%></a>
                    </td>
                    <td align="center">
                        <%#DealName(Eval("REALNAME"), Eval("NICKNAME"))%>
                    </td>
                    <td align="center">
                        <%#Eval("TotalGrade")%>
                    </td>
                    <td align="center">
                        <%#Eval("Counting")%>
                    </td>
                    <td align="center">
                        <%#Eval("useenergy")%>
                    </td>
                    <td align="center">
                        <%#Eval("Energy")%>
                    </td>
                    <td align="center">
                        <%#Eval("GetEnergy")%>
                    </td>
               </tr>
           </ItemTemplate>
           <FooterTemplate>
               </table>
           </FooterTemplate>
        </asp:Repeater>
      </div>
      </td>
      </tr>
      
      </table>
      </div>
</form>
</body>
</html>
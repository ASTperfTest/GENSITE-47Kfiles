<%@ Page Language="C#" AutoEventWireup="true" CodeFile="UserGameLog.aspx.cs" Inherits="kmactivity_kmwebpuzzle_UserGameLog" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
     <link href="css/style.css" rel="stylesheet" type="text/css" />
      <script type="text/javascript" >
          function GoTop() {
              window.location.href = "/kmactivity/kmwebpuzzle/BackstageGameRank.aspx?a=edrftg";
          }
    
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
             <td class="content">
    <div class="Page" style="padding-left:20px; padding-top:20px;"><br />
    <input type="button" value="回排行榜" onclick="GoTop()" /> 
    <br/>
    帳號：<asp:Label ID="UserName" runat="server" Text=""></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
    姓名/暱稱：<asp:Label ID="NickName" runat="server" Text=""></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
     E-mail：<asp:Label ID="UserMail" runat="server" Text=""></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
     <br/>
     尚未使用點數:<asp:Label ID="LblEnergy" runat="server" Text=""></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
     已使用點數:<asp:Label ID="LblUseEnergy" runat="server" Text=""></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
     總點數:<asp:Label ID="LblAllEnergy" runat="server" Text=""></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
    <br/>
    <br/>
    <asp:HyperLink ID="linkExport"  runat="server" Target="_blank">匯出報表</asp:HyperLink> &nbsp;&nbsp;&nbsp;  
     <br/>
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
                <table width="50%" class="type02" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <th>
                        &nbsp;
                        </th>
                        <th>
                            日期
                        </th>
                        <th>
                            完成/放棄
                        </th>
                        <th>
                            難度
                        </th>
                        <th>
                            活動得分
                        </th>
                        <th>
                            使用點數
                        </th>
                        <th>
                            縮圖
                        </th>
                    </tr>
            </HeaderTemplate>
            <ItemTemplate>
                <tr>
                    <td align="center">
                        <%#Eval("Row") %>
                    </td>   
                    <td align="center">
                        <%#Eval("gametime")%>
                    </td>
                    <td align="center">
                        <%#GetStates(Eval("picstate"))%>
                    </td>
                     <td align="center">
                        <%#GetDifficult(Eval("difficult"))%>
                    </td>
                    <td align="center">
                        <%#GetScore(Eval("difficult"), Eval("picstate"))%>
                    </td>
                    <td align="center">
                         <%#Eval("useenergy")%>
                    </td>
                    <td align="center">
                        <img src="/kmactivity/puzzle/puzzlePics/<%#Eval("pic_no")%>/<%#Eval("pic_name")%>.jpg" width="30px" height="30px" /> 
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
    </form>
</body>
</html>

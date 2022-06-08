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
<!--�����˦��j�P�W�w�gOK�F-->
    <form id="form1" runat="server">
    <div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
             <td class="content">
             <div class="content_mid">
        <div class="Page">
			<br/>
		   �z�����G<br/>
		   
		   <table><tr>
		                <td>�|��:
		                </td>
		                <td><asp:TextBox ID="TextBoxMember" runat="server" Width="100px"></asp:TextBox>&nbsp;(�b���B�m�W�B�ʺ�)
		                </td>
		          </tr>
                  
		          <tr>
		                <td>����:
		                </td>
		                <td><asp:TextBox ID="TextScoreLowerBound" runat="server" Width="30px"></asp:TextBox>~<asp:TextBox ID="TextScoreUpperBound" runat="server" Width="30px"></asp:TextBox>
		                <asp:RegularExpressionValidator id="RegularExpressionValidator1" ValidationExpression="\d*" runat="server" ControlToValidate="TextScoreUpperBound" ErrorMessage="��o�_���ƥu���J�Ʀr" Display="None" SetFocusOnError="true"/>
		                <asp:RegularExpressionValidator id="RegularExpressionValidator2" ValidationExpression="\d*" runat="server" ControlToValidate="TextScoreLowerBound" ErrorMessage="��o�_���ƥu���J�Ʀr" Display="None" SetFocusOnError="true"/>
		                </td>
		          </tr>
		          <tr>
		                <td>������:
		                </td>
		                <td>
                        <asp:TextBox ID="TextBox1" runat="server" Width="30px"></asp:TextBox>~<asp:TextBox ID="TextBox2" runat="server" Width="30px"></asp:TextBox>
		                </td>
		          </tr>
		          <tr>
		                <td>����϶�:
		                </td>
		                <td>
		                <asp:TextBox ID="TxtStartDate" runat="server" Width="80px" ></asp:TextBox>~<asp:TextBox ID="TxtEndDate" runat="server" Width="80px"></asp:TextBox>
		                
		                </td>
		          </tr>

		          <tr>
		                <td colspan="2">
                            <asp:Button ID="Button1" runat="server" Text="�d��" onclick="Button1_Click" />&nbsp;&nbsp;
		                </td>
		          </tr>
		   </table>
            <asp:HyperLink ID="linkExport"  runat="server" Target="_blank">�ץX����</asp:HyperLink> &nbsp;&nbsp;&nbsp;  
            <asp:Button ID="Button2" runat="server" onclick="Button2_Click" Text="���ܨϥΪ̪��A" /> &nbsp;&nbsp;&nbsp;  
            <asp:LinkButton ID="linkCheckCheat" runat="server" OnClick="CheckCheat">���ҿ��~�M��</asp:LinkButton>
		   </div>
           
        ��<asp:Label ID="PageNumberText" runat="server" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server"  CssClass="Number"/>���A
   	    �@<asp:Label ID="TotalRecordText" runat="server" CssClass="Number" />����ơA     
        <asp:LinkButton ID="preLink" runat="server" Visible="false" OnClick="preLinkAct">
            <asp:Label ID="preText" runat="server" Text="�W�@��" />&nbsp;
        </asp:LinkButton>
        ���ܲ�<asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" OnSelectedIndexChanged="ChangePageNumber" runat="server" /> �� &nbsp;
        <asp:LinkButton ID="nextLink" runat="server" Visible="false" OnClick="nextLinkAct">
            <asp:Label ID="nextText" runat="server" Text="�U�@��" />
        </asp:LinkButton>
        �A�C��                      
        <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" OnSelectedIndexChanged="ChangePageSize" runat="server">
        <asp:ListItem Selected="True" Value="15">15</asp:ListItem>                    
        <asp:ListItem Value="30">30</asp:ListItem>
        <asp:ListItem Value="50">50</asp:ListItem>
        </asp:DropDownList>�����

        <asp:Repeater ID="rpList" runat="server">
            <HeaderTemplate>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                        <th>
                         &nbsp;
                        </th>
                        <th>
                            ���v
                        </th>
                        <th>
                            �b��
                        </th>
                        <th>
                            �m�W�x�ʺ�
                        </th>
                        <th>
                            ���ʱo��
                        </th>
                        <th>
                            ������
                        </th>
                        <th>
                            �ϥ��I��
                        </th>
                        <th>
                            ���ϥ��I��
                        </th>
                        <th>
                            �`�I��
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
<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PhoneticNotation.aspx.cs" Inherits="jigsaw10_PhoneticNotation"  %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link type="text/css" rel="stylesheet" href="../css/list.css" />
    <link type="text/css" rel="stylesheet" href="../css/layout.css" />
    <link type="text/css" rel="stylesheet" href="../css/setstyle.css" />
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div id="FuncName">
            <h1>
                罕見字注音管理</h1>
            <div id="Nav">
            <a href="Index.aspx" >回前頁</a>
                
            </div>
            <div id="ClearFloat">
            </div>
        </div>
        <div id="FormName">
        </div>
    <div>
    <table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
            
            <tr>
                <td width="95%" colspan="2" height="230" valign="top">
                    <div align="center" style="width:60%">
                        <asp:Repeater ID="rptList" runat="server" OnItemCommand="Repeater_OnItemCommand" >                        
                            <HeaderTemplate>
                                <table id="ListTable" width="100%" cellpadding="0" cellspacing="1">
                                    <tr align="left">
                                        <th style="width:20%">
                                            
                                        </th>
                                        <th >
                                            字
                                        </th>
                                        <th >
                                            音
                                        </th>
                                    </tr>
                            </HeaderTemplate>
                            <ItemTemplate>
                                <tr>
                                    <td>
                                    [<asp:LinkButton ID="btndelete" runat="server" Text="刪除" CommandName="Delete" CommandArgument='<%#Eval("word") %>'></asp:LinkButton>]&nbsp;
                                    [<asp:LinkButton ID="btUpdate" runat="server" Text="修改" CommandName="Update" CommandArgument='<%#Eval("word") %>' ></asp:LinkButton>]
                                    </td>
                                    <td>
                                        <%#Eval("word") %>
                                    </td>
                                    <td>
                                        <%# Eval("PhoneticNotation1")%>
                                    </td>
                                </tr>
                            </ItemTemplate>
                            <FooterTemplate>
                                </table>
                                <div align="center">
                                
                                </div>
                                    
                            </FooterTemplate>
                        </asp:Repeater>
                    </div>
                </td>
            </tr>
        </table>
        <table id="ListTable">
                                    <tr>
                                        <th style="width:15%" align="right">
                                            字：
                                        </th>
                                        <td class="eTableContent">
                                            <asp:TextBox ID="txtWord" runat="server"></asp:TextBox>
                                        </td>                                        
                                    </tr><tr>
                                        <th style="width:15%" align="right">
                                            音：
                                        </th>
                                        <td class="eTableContent">
                                            <asp:TextBox ID="txtphonetec" runat="server"></asp:TextBox>
                                            <asp:Button ID="bsave" runat="server" Text="新增" onclick="bsave_Click"  />
                                            <asp:Button ID="bUpdate" runat="server" Text="確定" onclick="bUpdate_Click" />
                                            <asp:Button ID="bCancel" runat="server" Text="取消" onclick="bCancel_Click" />
                                        </td>                                        
                                    </tr>
                                </table>
    </div>
    </form>
</body>
</html>

<%@ Page Language="C#" AutoEventWireup="true" CodeFile="bgMusicList.aspx.cs" Inherits="bgMusic_bgMusicList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link rel="stylesheet" href="/inc/setstyle.css" />
    <link href="../css/list.css" rel="stylesheet" type="text/css" />
    <link href="../css/layout.css" rel="stylesheet" type="text/css" />
    <title>背景音樂管理</title>
</head>
<body>
    <form id="form1" runat="server">
    <div id="FuncName">
        <h1>
            主題館／圖片播放版背景音樂管理</h1>
        <div id="Nav">
        <a href="bgMusicEdit.aspx?type=Add" title="新增">新增</a><a href="../index.aspx">回前頁</a>
        </div>
        <div id="ClearFloat">
        </div>
    </div>
    <div id="FormName">
    </div>
    <div>
        <table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
            <tr>
                <td>
                    <div align="center">
                        <asp:Repeater ID="rptList" runat="server" OnItemCommand="Repeater_OnItemCommand">
                        <HeaderTemplate>
                                <table cellspacing="0" id="ListTable">
                                    <tr>
                                        <th style="width:150px">
                                            
                                        </th>
										<th >
                                            曲目標題
                                        </th>
                                        <th >
                                            檔名
                                        </th>                                        
                                    </tr>
                            </HeaderTemplate>
                            <ItemTemplate>
                                <tr>
                                    <td  align="center">
                                        <asp:Button ID="Edit" runat="server" CommandName="Edit" CommandArgument='<%# Eval("bgMusicID") %>' Text="編輯" />
                                        <asp:Button ID="btnDelete" runat="server" OnClientClick="javascript:return(confirm('確定要刪除?'));"
                                            Text="刪除" CommandName="Delete" CommandArgument='<%# Eval("bgMusicID")%>' />
                                    </td>
                                    <td><%#Eval("Title")%>
                                    </td>
                                    <td><%#Eval("FileName") %>;
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
    <asp:SqlDataSource ID="DSBgMusic" runat="server"></asp:SqlDataSource>
    </form>
</body>
</html>

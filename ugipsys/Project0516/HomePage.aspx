<%@ Page Language="C#" AutoEventWireup="true" CodeFile="HomePage.aspx.cs" Inherits="HomePage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <link href="css/list.css" rel="stylesheet" type="text/css" />
    <link href="css/layout.css" rel="stylesheet" type="text/css" />
    <link href="css/theme.css" rel="stylesheet" type="text/css" />
    <link href="css/Display.css" rel="stylesheet" type="text/css" />
    <title>主題館管理</title>
    
    
</head>
<body>

<div id="FuncName">
	<h1>新增主題館</h1>
	<div id="ClearFloat"></div>
</div>

<div class="step">步驟： 1. 填寫基本資料 > 2. 設定版面風格 > 3. 導覽架構設定 > <span>4. 首頁配置</span></div>

<form id="form1" runat="server">
<!--列表-->
    <table cellspacing="0" class="setting">
    <caption>【設定主題館首頁配置】(請依據版面配置方式，選取各編號區塊的顯示方式)</caption>
    
    <tr>
        <td>
            <h5>您選取的版面配置 》</h5>
            <asp:ImageButton ID="LayoutPic" runat="server" />    
        </td>
                
        <td>
            <h5>區塊顯示資料＆顯示方式設定 》</h5>
            <!--Block1-->
            <p>區塊『<strong><em> 1 </em></strong>』要顯示的資料：
                <asp:DropDownList ID="Block1DataNode" runat="server" CssClass="box" DataSourceID="ObjectDataSourceItem" DataTextField="catName" DataValueField="ctNodeId">
                    <asp:ListItem>1</asp:ListItem>
                    <asp:ListItem>2</asp:ListItem>
                </asp:DropDownList>
                <asp:Button ID="Block1Down" runat="server" Text="↓" CssClass="button" OnClick="Block1Down_Click" />
            </p>
            <p>
                顯示標題：<asp:TextBox ID="Block1Title" runat="server" CssClass="box" MaxLength="8" Width="100px"></asp:TextBox>(請勿超過8個字)
            </p>            
            <fieldset>
                <legend> < 顯示方式 > </legend>
                <p style="width: 430px">顯示資料欄位：<br />
                <asp:CheckBox ID="Block1IsTitle" runat="server" Checked="True" Enabled="False"  Text="標題" />
                <asp:CheckBox ID="Block1IsPic" runat="server" Text="圖片" />
                <asp:CheckBox ID="Block1IsPostDate" runat="server" Text="張貼日期" />
                 <asp:CheckBox ID="Block1IsExcerpt" runat="server" Text="摘要(字數：" />
                 <asp:TextBox ID="Block1ContentLength" runat="server" Width="20px" CssClass="box"></asp:TextBox>字)
                </p>
                <p>
                              
                 <div class="center">                            
                    <asp:RadioButton ID="Block1Type1" runat="server" GroupName="Block1MakeUp" Text="(用於顯示多筆)" Checked="True" /><br />
                    <asp:ImageButton ID="Block1ShowType1" runat="server" Width="120" ImageUrl="~/images/lp01.gif" />
                 </div>
                 <div class="center">     
                    <asp:RadioButton ID="Block1Type2" runat="server" GroupName="Block1MakeUp" Text="(用於顯示多筆)" /><br />
                    <asp:ImageButton ID="Block1ShowType2" runat="server" Width="120" ImageUrl="~/images/lp03.gif" />
                 </div>
                 <div class="center">                        
                    <asp:RadioButton ID="Block1Type3" runat="server" GroupName="Block1MakeUp" Text="(用於顯示<em>1</em>筆)"/><br />
                    <asp:ImageButton ID="Block1ShowType3" runat="server" Width="120" ImageUrl="~/images/homelist.jpg" />
                 </div>
                 </p>
                 <p>             
                     <asp:Panel ID="Block1PanelSqlTop" runat="server">
                     顯示資料數量：<asp:TextBox ID="Block1SqlTop" runat="server" Width="20px" CssClass="box"></asp:TextBox>筆                 
                     </asp:Panel>
                 </p>
                
            </fieldset>            
            <br />
            <!--End Block1-->
            <!--Block2-->
            <p>區塊『<strong><em> 2 </em></strong>』要顯示的資料：
                <asp:DropDownList ID="Block2DataNode" runat="server" CssClass="box" DataSourceID="ObjectDataSourceItem" DataTextField="catName" DataValueField="ctNodeID">
                </asp:DropDownList>
                <asp:Button ID="Block2Up" runat="server" Text="↑" CssClass="button" OnClick="Block1Down_Click"/>
                <asp:Button ID="Block2Down" runat="server" Text="↓" CssClass="button" OnClick="Block2Down_Click" />
            </p>
            <p>
                顯示標題：<asp:TextBox ID="Block2Title" runat="server" CssClass="box" MaxLength="8" Width="100px"></asp:TextBox>(請勿超過8個字)
            </p>  
            <fieldset>
                <legend> < 顯示方式 > </legend>
                <p>顯示資料欄位：<br />
                 <asp:CheckBox ID="Block2IsTitle" runat="server" Checked="True" Enabled="False"  Text="標題" />
                 <asp:CheckBox ID="Block2IsPic" runat="server" Text="圖片" />
                 <asp:CheckBox ID="Block2IsPostDate" runat="server" Text="張貼日期" />
                 <asp:CheckBox ID="Block2IsExcerpt" runat="server" Text="摘要(字數：" />
                 <asp:TextBox ID="Block2ContentLength" runat="server" Width="20px" CssClass="box"></asp:TextBox>字)                
                </p>
                <p>
                </p>    
                 <div class="center">                    
                    <asp:RadioButton ID="Block2Type1" runat="server" GroupName="Block2MakeUp" Text="(用於顯示多筆)" Checked="True" /><br />
                     <asp:ImageButton ID="Block2ShowType1" runat="server" Width="120" ImageUrl="~/images/lp01.gif" />
                      </div>
                      <div class="center">
                    <asp:RadioButton ID="Block2Type2" runat="server" GroupName="Block2MakeUp" Text="(用於顯示多筆)"/><br />
                    <asp:ImageButton ID="Block2ShowType2" runat="server" Width="120" ImageUrl="~/images/lp03.gif" />
                     </div>
                      <div class="center">       
                    <asp:RadioButton ID="Block2Type3" runat="server" GroupName="Block2MakeUp" Text="(用於顯示<em>1</em>筆)"/><br />                          
                    <asp:ImageButton ID="Block2ShowType3" runat="server" Width="120" ImageUrl="~/images/homelist.jpg" />
                  </div>
                <p>                 
                    <asp:Panel ID="Block2PanelSqlTop" runat="server">
                     顯示資料數量：<asp:TextBox ID="Block2SqlTop" runat="server" Width="20px" CssClass="box"></asp:TextBox>筆                 
                    </asp:Panel>
                </p>
            </fieldset>
            <br />
            <!--End Block2-->
            
            <!--Block3-->
            <p>區塊『<strong><em> 3 </em></strong>』要顯示的資料：
                <asp:DropDownList ID="Block3DataNode" runat="server" CssClass="box" DataSourceID="ObjectDataSourceItem" DataTextField="catName" DataValueField="ctNodeID">
                </asp:DropDownList>
                <asp:Button ID="Block3Up" runat="server" Text="↑" CssClass="button" OnClick="Block2Down_Click"/>
                <asp:Button ID="Block3Down" runat="server" Text="↓" CssClass="button" OnClick="Block3Down_Click" />
            </p>
            <p>
                顯示標題：<asp:TextBox ID="Block3Title" runat="server" CssClass="box" MaxLength="8" Width="100px"></asp:TextBox>(請勿超過8個字)
            </p>  
            <fieldset>
                <legend> < 顯示方式 > </legend>
                <p style="width: 430px">顯示資料欄位：<br />
                 <asp:CheckBox ID="Block3IsTitle" runat="server" Checked="True" Enabled="False"  Text="標題" />
                 <asp:CheckBox ID="Block3IsPic" runat="server" Text="圖片" />
                 <asp:CheckBox ID="Block3IsPostDate" runat="server" Text="張貼日期" />
                 <asp:CheckBox ID="Block3IsExcerpt" runat="server" Text="摘要(字數：" />
                 <asp:TextBox ID="Block3ContentLength" runat="server" Width="20px" CssClass="box"></asp:TextBox>字)                
                </p>
                <p>
                
                </p>
                  <div class="center">  
                           <asp:RadioButton ID="Block3Type1" runat="server" GroupName="Block3MakeUp" Text="(用於顯示多筆)" Checked="True" /><br />     
                            <asp:ImageButton ID="Block3ShowType1" runat="server" Width="120" ImageUrl="~/images/lp01.gif" />
                  </div>
                   <div class="center">  
                              <asp:RadioButton ID="Block3Type2" runat="server" GroupName="Block3MakeUp" Text="(用於顯示多筆)"  /><br />  
                               <asp:ImageButton ID="Block3ShowType2" runat="server" Width="120" ImageUrl="~/images/lp03.gif" />
                  </div>
                   <div class="center">  
                            <asp:RadioButton ID="Block3Type3" runat="server" GroupName="Block3MakeUp" Text="(用於顯示<em>1</em>筆)" /><br />    
                            <asp:ImageButton ID="Block3ShowType3" runat="server" Width="120" ImageUrl="~/images/homelist.jpg" />
                  </div>
                     
                 <p>    
                     <asp:Panel ID="Block3PanelSqlTop" runat="server" Height="50px" Width="125px">
                     顯示資料數量：<asp:TextBox ID="Block3SqlTop" runat="server" Width="20px" CssClass="box"></asp:TextBox>筆
                     </asp:Panel>
                 </p>         
            </fieldset>
            <br />
            <!--End Block3-->
            
            <!--Block4-->
            <p>區塊『<strong><em> 4 </em></strong>』要顯示的資料：
                <asp:DropDownList ID="Block4DataNode" runat="server" CssClass="box" DataSourceID="ObjectDataSourceItem" DataTextField="catName" DataValueField="ctNodeID">
                </asp:DropDownList>
                <asp:Button ID="Block4Up" runat="server" Text="↑" CssClass="button" OnClick="Block3Down_Click"/>
                <asp:Button ID="Block4Down" runat="server" Text="↓" CssClass="button" OnClick="Block4Down_Click" />
            </p>
            <p>
                顯示標題：<asp:TextBox ID="Block4Title" runat="server" CssClass="box"  MaxLength="8" Width="100px"></asp:TextBox>(請勿超過8個字)
            </p>  
            <fieldset>
                <legend> < 顯示方式 > </legend>
                <p>顯示資料欄位：<br />
                 <asp:CheckBox ID="Block4IsTitle" runat="server" Checked="True" Enabled="False"  Text="標題" />
                 <asp:CheckBox ID="Block4IsPic" runat="server" Text="圖片" />
                 <asp:CheckBox ID="Block4IsPostDate" runat="server" Text="張貼日期" />
                 <asp:CheckBox ID="Block4IsExcerpt" runat="server" Text="摘要(字數：" />
                 <asp:TextBox ID="Block4ContentLength" runat="server" Width="20px" CssClass="box"></asp:TextBox>字)                
                </p>
                <p>
                  </p>
                 <div class="center">
                        <asp:RadioButton ID="Block4Type1" runat="server" GroupName="Block4MakeUp" Text="(用於顯示多筆)" Checked="True" /><br />
                        <asp:ImageButton ID="Block4ShowType1" runat="server" Width="120" ImageUrl="~/images/lp01.gif" />
                 </div>
                 <div class="center">
                        <asp:RadioButton ID="Block4Type2" runat="server" GroupName="Block4MakeUp" Text="(用於顯示多筆)" /><br />
                        <asp:ImageButton ID="Block4ShowType2" runat="server" Width="120" ImageUrl="~/images/lp03.gif" />
                 </div>
                 <div class="center">
                        <asp:RadioButton ID="Block4Type3" runat="server" GroupName="Block4MakeUp" Text="(用於顯示<em>1</em>筆)"  /><br />
                     <asp:ImageButton ID="Block4ShowType3" runat="server" Width="120" ImageUrl="~/images/homelist.jpg" />
                 
                 </div>
                  <p>
                      <asp:Panel ID="Block4PanelSqlTop" runat="server">
                      顯示資料數量：<asp:TextBox ID="Block4SqlTop" runat="server" Width="20px" CssClass="box"></asp:TextBox>筆</asp:Panel>
                  </p>                         
            </fieldset>
            <br />
            <!--End Block4-->
        
            <!--Block5-->
            <p>區塊『<strong><em> 5 </em></strong>』要顯示的資料：
                <asp:DropDownList ID="Block5DataNode" runat="server" CssClass="box" DataSourceID="ObjectDataSourceItem" DataTextField="catName" DataValueField="ctNodeID">
                </asp:DropDownList>
                <asp:Button ID="Block5Up" runat="server" Text="↑" CssClass="button" OnClick="Block4Down_Click"/>&nbsp;
            </p>
            <p>
                顯示標題：<asp:TextBox ID="Block5Title" runat="server" CssClass="box" MaxLength="8" Width="100px"></asp:TextBox>(請勿超過8個字)
            </p>  
            <fieldset>
                <legend> < 顯示方式 > </legend>
                <p>顯示資料欄位：<br />
                 <asp:CheckBox ID="Block5IsTitle" runat="server" Checked="True" Enabled="False"  Text="標題" />
                 <asp:CheckBox ID="Block5IsPic" runat="server" Text="圖片" />
                 <asp:CheckBox ID="Block5IsPostDate" runat="server" Text="張貼日期" />
                 <asp:CheckBox ID="Block5IsExcerpt" runat="server" Text="摘要(字數：" />
                 <asp:TextBox ID="Block5ContentLength" runat="server" Width="20px" CssClass="box"></asp:TextBox>字)                
                </p>
                <p>
                 
                </p>
                 <div class="center">
                      <asp:RadioButton ID="Block5Type1" runat="server" GroupName="Block5MakeUp" Text="(用於顯示多筆)" Checked="True" /><br />  
                      <asp:ImageButton ID="Block5ShowType1" runat="server" Width="120" ImageUrl="~/images/lp01.gif" />
                 </div>
                 <div class="center">
                      <asp:RadioButton ID="Block5Type2" runat="server" GroupName="Block5MakeUp" Text="(用於顯示多筆)"/><br /> 
                      <asp:ImageButton ID="Block5ShowType2" runat="server" Width="120" ImageUrl="~/images/lp03.gif" /> 
                 </div>
                 <div class="center">
                      <asp:RadioButton ID="Block5Type3" runat="server" GroupName="Block5MakeUp" Text="(用於顯示<em>1</em>筆)"/><br />
                       <asp:ImageButton ID="Block5ShowType3" runat="server" Width="120" ImageUrl="~/images/homelist.jpg" />
                 </div>
                <p>
                    <asp:Panel ID="Block5PanelSqlTop" runat="server" Height="50px" Width="125px">
                     顯示資料數量：<asp:TextBox ID="Block5SqlTop" runat="server" Width="20px" CssClass="box"></asp:TextBox>筆
                    </asp:Panel>
                </p>
            </fieldset>
            <br />
            <!--End Block5-->        
        </td>
    </tr>
    </table>
    
    <div class="settingbutton">
        <asp:Button ID="Back" runat="server" Text="上一步" OnClick="Back_Click" />
        <asp:Button ID="Next" runat="server" Text="下一步" OnClick="Next_Click" />
        <asp:Button ID="Save" runat="server" Text="儲存設定" OnClick="Save_Click" />
        <asp:Button ID="Home" runat="server" Text="取消(回首頁)" OnClick="Home_Click" />
</div>

    <div>
        <asp:ObjectDataSource ID="ObjectDataSourceItem" runat="server" SelectMethod="NodeItem"
            TypeName="HomePageSql">
            <SelectParameters>
                <asp:SessionParameter Name="TreeRootId" SessionField="TreeRootId" Type="Int32" />
            </SelectParameters>
        </asp:ObjectDataSource>
        &nbsp;
    
    
    </div>
</form>
                        
</body>
</html>

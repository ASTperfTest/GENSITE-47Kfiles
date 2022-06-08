<%@ Control Language="C#" AutoEventWireup="true" CodeFile="BlockControl.ascx.cs" Inherits="BlockControl" %>

 <!--Block1-->
            <p>區塊『<strong><em> 
                <asp:Label ID="Block1Num" runat="server" Text="1"></asp:Label> </em></strong>』要顯示的資料：
                <asp:DropDownList ID="Block1DataNode" runat="server" CssClass="box">                   
                </asp:DropDownList>
                <asp:Button ID="Block1Up" runat="server" Text="↑" CssClass="button" OnClick="Block1Up_Click"/>
                <asp:Button ID="Block1Down" runat="server" Text="↓" CssClass="button" OnClick="Block1Down_Click" />
            </p>
            <p>
                顯示標題：<asp:TextBox ID="Block1Title" runat="server" CssClass="box" MaxLength="8" Width="100px"></asp:TextBox>(請勿超過8個字)
            </p>            
            <fieldset>
                <legend> < 顯示方式 > </legend>
                <p>顯示資料欄位：<br />
                <asp:CheckBox ID="Block1IsTitle" runat="server" Checked="True" Enabled="False"  Text="標題" />
                <asp:CheckBox ID="Block1IsPic" runat="server" Text="圖片" />
                <asp:CheckBox ID="Block1IsPostDate" runat="server" Text="張貼日期" />
                 <asp:CheckBox ID="Block1IsExcerpt" runat="server" Text="摘要(字數：" />
                 <asp:TextBox ID="Block1ContentLength" runat="server" Width="20px" CssClass="box"></asp:TextBox>字)
                 </p>
                <p>
                              
                 <div class="center">                            
                    <asp:RadioButton ID="Block1Type1" runat="server" GroupName="Block1MakeUp" Text="(用於顯示多筆)" Checked="True" /><br />
                    <asp:Image ID="Block1ShowType1" runat="server" Width="120" ImageUrl="~/images/lp01.gif" /></div>
                 <div class="center">     
                    <asp:RadioButton ID="Block1Type2" runat="server" GroupName="Block1MakeUp" Text="(用於顯示多筆)" /><br />
                    <asp:Image ID="Block1ShowType2" runat="server" Width="120" ImageUrl="~/images/lp03.gif" />
                 </div>
                 <div class="center">                        
                    <asp:RadioButton ID="Block1Type3" runat="server" GroupName="Block1MakeUp" Text="(用於顯示<em>1</em>筆)"/><br />
                    <asp:Image ID="Block1ShowType3" runat="server" Width="120" ImageUrl="~/images/homelist.jpg" />
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
                 

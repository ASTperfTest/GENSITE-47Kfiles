<%@ Page Title="" Language="C#" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="true" CodeFile="Detail.aspx.cs" Inherits="Detail" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <script src="Scripts/swfobject_modified.js" type="text/javascript"></script>
    
    <div class="path" style="float: left; padding-top: 10px">
        目前位置：
    </div>
    <div style="float: left; width: 80%">
        <asp:Literal ID="nav" runat="server"></asp:Literal>
    </div>
    <div class="jigsaw">
        <div class="head">
        </div>
        <div class="body">
            <div id="crop">
                <asp:Repeater ID="rptList10" runat="server">
                    <ItemTemplate>
                        <h5 style="font-weight: bold; font-size: 12pt; line-height: 21pt; color: #060; background: #dde89a;
                            border-bottom: #ccc dashed 1px; margin: 0 0 18px 0; padding: 0 6px;">
                            <%# Eval("sTitle")%></h5>

                        <%--Added by Leo    2011-06-20      加入圖片輪播:jQuery Supersized    Start--%>                       
                        <center><%# GetImgString(Eval("sTitle"), Eval("xImgFile"))%></center></center>
                        
                        <%--Added by Leo    2011-06-20      加入圖片輪播:jQuery Supersized     End --%>
                        <br />
                        <%# (string)Eval("alias") == "" ? "" : "<p>別名：" + Eval("alias") + "</p>"%>
                        <%# (string)Eval("description") == "" ? "" : "<p>簡介：" + Eval("description") + "</p>"%>
                        <p>
                            產期：<%# Eval("season")%></p>
                        <p>
                            產地：<%# Eval("origin")%></p>
                        <%# (string)Eval("feature") == "" ? "" : "<p>特性：" + Eval("feature") + "</p>"%>
                        <%# (string)Eval("nutritionValue") == "" ? "" : "<p>營養價值：" + Eval("nutritionValue") + "</p>"%>
                        <%# (string)Eval("selectionMethod") == "" ? "" : "<p>挑選方式：" + Eval("selectionMethod") + "</p>"%>
                        <%# (string)Eval("nutrient") == "" ? "" : "<p>營養成分：" + Eval("nutrient") + "</p>"%>
                        <%# (string)Eval("cuisine") == "" ? "" : "<p>料理：" + Eval("cuisine") + "</p>"%>
                        <%# (string)Eval("variety") == "" ? "" : "<p>品種：" + Eval("variety") + "</p>"%>
                        <%# (string)Eval("note") == "" ? "" : "<p>備註：" + Eval("note") + "</p>"%>
                        <%# (string)Eval("tips") == "" ? "" : "<p>小秘訣：" + Eval("tips") + "</p>"%>
                        <p align="right">
                            【資料來源:農糧署】</p>
                    </ItemTemplate>
                </asp:Repeater>
                <asp:Repeater ID="rptList11" runat="server">
                    <ItemTemplate>
                        <h5 style="font-weight: bold; font-size: 12pt; line-height: 21pt; color: #060; background: #dde89a;
                            border-bottom: #ccc dashed 1px; margin: 0 0 18px 0; padding: 0 6px;">
                            <%# PhoneticNotation.GetNotation(Eval("sTitle").ToString())%></h5>
                            <%--Added by Leo    2011-06-20      加入圖片輪播:jQuery Supersized    Start--%>
                            <center><%# GetImgString(Eval("sTitle"), Eval("xImgFile"))%></center></center>
                            
                            學名：<%# Eval("scientificName")%></p>
                        <%# (string)Eval("family") == "" ? "" : "<p>科名：" + Eval("family") + "</p>"%>
                        <%# (string)Eval("commonName") == "" ? "" : "<p>科俗名：" + Eval("commonName") + "</p>"%>
                        <p>
                            產期：<%# Eval("season")%></p>
                        <p>
                            產地：<%# Eval("origin")%></p>
                        <%# (string)Eval("distributionInTaiwan") == "" ? "" : "<p>地理分布：" + Eval("distributionInTaiwan") + "</p>"%>
                        <%# (string)Eval("habitats") == "" ? "" : "<p>棲息環境：" + Eval("habitats") + "</p>"%>
                        <%# (string)Eval("reference") == "" ? "" : "<p>參考文獻：" + Eval("reference") + "</p>"%>
                        <%# (string)Eval("characteristic") == "" ? "" : "<p>形態特徵：" + Eval("characteristic") + "</p>"%>
                        <%# (string)Eval("habitatsType") == "" ? "" : "<p>棲所生態：" + Eval("habitatsType") + "</p>"%>
                        <%# (string)Eval("distribution") == "" ? "" : "<p>地理分布：" + Eval("distribution") + "</p>"%>
                        <%# (string)Eval("utility") == "" ? "" : "<p>漁業利用：" + Eval("utility") + "</p>"%>
                        <p align="right">
                            【資料來源:<a href='http://fishdb.sinica.edu.tw/' target='_blank'>台灣魚類資料庫</a>】</p>
                    </ItemTemplate>
                </asp:Repeater>
            </div>
            <asp:Repeater ID="rptList20" runat="server">
                <HeaderTemplate>
                    <div class="jigsawnew">
                        <h5>
                            最新議題</h5>
                </HeaderTemplate>
            </asp:Repeater>
            <asp:Repeater ID="rptList21" runat="server">
                <HeaderTemplate>
                    <div class="content">
                </HeaderTemplate>
                <ItemTemplate>
                    <%# Eval("xImgFile")%><h6>
                        <a target="_nwGip" href="<%# Eval("path")%>">
                            <%# Eval("sTitle")%></a></h6>
                    <p>
                        <%# Eval("xBody")%>...<a target="_nwGip" href="<%# Eval("path")%>"> 詳全文 </a>
                    </p>
                </ItemTemplate>
                <FooterTemplate>
                    </div></FooterTemplate>
            </asp:Repeater>
            <asp:Repeater ID="rptList22" runat="server">
                <HeaderTemplate>
                    <ol>
                </HeaderTemplate>
                <ItemTemplate>
                    <li><a target="_nwGip" href="<%# Eval("path")%>">
                        <%# Eval("sTitle")%></a></li>
                </ItemTemplate>
                <FooterTemplate>
                    </ol></FooterTemplate>
            </asp:Repeater>
            <asp:Repeater ID="rptList29" runat="server">
                <HeaderTemplate>
                    </div>
                </HeaderTemplate>
            </asp:Repeater>
            <asp:Repeater ID="rptList30" runat="server">
                <HeaderTemplate>
                    <h5>
                        議題關聯知識文章</h5>
                </HeaderTemplate>
            </asp:Repeater>
            <asp:Repeater ID="rptList31" runat="server">
                <HeaderTemplate>
                    <table class="<%=reOrders[0].css%>" summary="結果條列式">
                        <tbody>
                            <tr>
                                <th colspan="3" scope="col">
                                    <%=reOrders[0].title%>
                                </th>
                            </tr>
                </HeaderTemplate>
                <ItemTemplate>
                    <tr>
                        <td align="center" width="5%">
                            <%# Eval("index")%>.
                        </td>
                        <td>
                            <a target="_nwGip" href="<%# Eval("path")%>">
                                <%# Eval("sTitle")%></a>
                        </td>
                    </tr>
                </ItemTemplate>
                <FooterTemplate>
                    </tbody></table></FooterTemplate>
            </asp:Repeater>
            <asp:Repeater ID="rptList32" runat="server">
                <HeaderTemplate>
                    <table class="<%=reOrders[1].css%>" summary="結果條列式">
                        <tbody>
                            <tr>
                                <th colspan="3" scope="col">
                                    <%=reOrders[1].title%>
                                </th>
                            </tr>
                </HeaderTemplate>
                <ItemTemplate>
                    <tr>
                        <td align="center" width="5%">
                            <%# Eval("index")%>.
                        </td>
                        <td>
                            <a target="_nwGip" href="<%# Eval("path")%>">
                                <%# Eval("sTitle")%></a>
                        </td>
                    </tr>
                </ItemTemplate>
                <FooterTemplate>
                    </tbody></table></FooterTemplate>
            </asp:Repeater>
            <asp:Repeater ID="rptList33" runat="server">
                <HeaderTemplate>
                    <table class="<%=reOrders[2].css%>" summary="結果條列式">
                        <tbody>
                            <tr>
                                <th colspan="3" scope="col">
                                    <%=reOrders[2].title%>
                                </th>
                            </tr>
                </HeaderTemplate>
                <ItemTemplate>
                    <tr>
                        <td align="center" width="5%">
                            <%# Eval("index")%>.
                        </td>
                        <td>
                            <a target="_nwGip" href="<%# Eval("path")%>">
                                <%# Eval("sTitle")%></a>
                        </td>
                    </tr>
                </ItemTemplate>
                <FooterTemplate>
                    </tbody></table></FooterTemplate>
            </asp:Repeater>
            <asp:Repeater ID="rptList34" runat="server">
                <HeaderTemplate>
                    <table class="<%=reOrders[3].css%>" summary="結果條列式">
                        <tbody>
                            <tr>
                                <th colspan="3" scope="col">
                                    <%=reOrders[3].title%>
                                </th>
                            </tr>
                </HeaderTemplate>
                <ItemTemplate>
                    <tr>
                        <td align="center" width="5%">
                            <%# Eval("index")%>.
                        </td>
                        <td>
                            <a target="_nwGip" href="<%# Eval("path")%>">
                                <%# Eval("sTitle")%></a>
                        </td>
                    </tr>
                </ItemTemplate>
                <FooterTemplate>
                    </tbody></table></FooterTemplate>
            </asp:Repeater>
            <asp:Repeater ID="rptList41" runat="server">
                <HeaderTemplate>
                    <table class="jigsawtype05" summary="結果條列式">
                        <tbody>
                            <tr>
                                <th colspan="3" scope="col">
                                    議題關聯影音
                                </th>
                            </tr>
                </HeaderTemplate>
                <ItemTemplate>
                    <tr>
                        <td align="center" width="5%">
                            <%# Eval("index")%>.
                        </td>
                        <td>
                            <a target="_nwGip" href="<%# Eval("path")%>">
                                <%# Eval("sTitle")%></a>
                        </td>
                    </tr>
                </ItemTemplate>
                <FooterTemplate>
                    </tbody></table></FooterTemplate>
            </asp:Repeater>
            <asp:Repeater ID="rptList51" runat="server">
                <HeaderTemplate>
                    <h5>
                        最佳資源推薦</h5>
                    <table class="jigsawtype03" summary="結果條列式">
                        <tbody>
                            <tr>
                                <th colspan="3" scope="col">
                                    推薦超聯結列表
                                </th>
                            </tr>
                </HeaderTemplate>
                <ItemTemplate>
                    <tr>
                        <td align="center" width="5%">
                            <%# Eval("index")%>.
                        </td>
                        <td>
                            <a target="_nwGip" href="<%# Eval("path")%>">
                                <%# Eval("sTitle")%></a>
                        </td>
                    </tr>
                </ItemTemplate>
                <FooterTemplate>
                    </tbody></table></FooterTemplate>
            </asp:Repeater>
            <h5>
                議題分享</h5>
            <table class="jigsawtype03" summary="結果條列式">
                <tbody>
                    <tr>
                        <th colspan="3" scope="col">
                            留言 | <span style="font-size: smaller;"><a href="javascript:toggleDiscussionSection();">我要分享</a></span>
                        </th>
                    </tr>
                    <asp:Repeater ID="rptList61" runat="server">
                        <ItemTemplate>
                            <tr>
                                <td width="5%" align="center">
                                    <%# Eval("index")%>.
                                </td>
                                <td width="15%" align="center" nowrap="nowrap">
                                    <%# Eval("iEditor")%> 發表於 <%# Eval("xpostdate")%>
                                </td>
                                <td>
                                    <%# Eval("xBody")%>
                                </td>
                            </tr>
                        </ItemTemplate>
                    </asp:Repeater>
                    <tr id="DiscussionBox" style="display: none">
                        <td colspan="3">
                            <div>
                                <textarea id="txtDisCussion" name="txtDisCussion" rows="10" cols="45"></textarea><br />
                                請輸入圖片驗證碼：
								
								<asp:Image ID="MemberCaptChaImage" runat="server" ImageUrl="" />
      <input type="hidden" value="<%=guid %>" id="MemberGuid" name="MemberGuid" /> 
      <label for="textfield">請輸入左邊圖片的數字</label><br />
    	<asp:TextBox ID="MemberCaptChaTBox" Columns="20" CssClass="txt" runat="server" />
								<input type="submit" value="確認" onclick="return checkNull();" /><input
                                            type="reset" value="重填" />
                            </div>
                        </td>
                    </tr>
                 </tbody>
            </table>
        </div>
        <div class="foot">
        </div>
        <div class="jigsawtop">
            <a href="#">top</a></div>
    </div>
    <script type="text/javascript" language="JavaScript">
        function toggleDiscussionSection() {
			var memberId = '<%=memberId%>';
            if (memberId == "") {
                alert('請先登入會員');
            }
			else {
			    var elem = document.getElementById('DiscussionBox');
				if (elem.style.display == "none")
				{ elem.style.display = ""; }
				else { elem.style.display = "none"; }			
			}
        }
        function checkNull() {
            var elem = document.getElementById('txtDisCussion').value;
            if (elem == "" || elem.trim() == "") {
                alert('請輸入留言內容');
                return false;
            }
            else
                return true;
        }
	</script>
</asp:Content>


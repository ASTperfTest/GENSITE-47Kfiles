<%@ Page Language="C#" MasterPageFile="~/4SideMasterPage.master"  AutoEventWireup="true" CodeFile="myspace.aspx.cs" Inherits="TreasureHunt_2011_myspace" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <link href="include/style.css" rel="stylesheet" type="text/css" />
    <link href="/js/greybox/gb_styles.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript">var GB_ROOT_DIR = "/js/greybox/";</script>
    <script type="text/javascript" src ="/js/greybox.js" ></script>
    <script type="text/javascript" src ="/js/greybox/AJS.js" ></script>
    <script type="text/javascript" src ="/js/greybox/AJS_fx.js" ></script>
    <script type="text/javascript" src ="/js/greybox/gb_scripts.js" ></script>
    <style type="text/css">
    span span.cti {
	position: absolute;
	top: -50px;
	* top: 0px;
	left: 0px;
	cursor:pointer;
	z-index:10;
    }
    .pointer 
    {
        cursor:pointer;
        border:1px solid #3366CC;
    }
    .trhdivover
    {
        background-color:red;
    }
    .trhdivchoose
    {
        border:1px solid red;
    }
    </style>
    <script type="text/javascript">
        var buttonDisable = "<%=buttonDisable %>";
        var voteDisplay = "<%=voteDisplay %>";
        var treasurebox;
        var changeflag = false;
        jQuery(document).ready(function () {
            if (buttonDisable == "false") {
                jQuery("#submitLotteryA").attr("disabled", false);
                jQuery("#submitLotteryB").attr("disabled", false);
                jQuery("#submitLotteryC").attr("disabled", false);
            }
            if (voteDisplay == "true") {
                jQuery("#present").show();
            }

            $('#<%=TreasureDropDownListClientId%>').change(function () {
                checkchange(this, "#<%=TreasureDropDownList2ClientId%>");
            });
            $('#<%=TreasureDropDownList2ClientId%>').change(function () {
                checkchange(this, "#<%=TreasureDropDownListClientId%>");
            });
            if (!treasurebox && $("#hitb").val() != undefined) {
                treasurebox = $.parseJSON(decodeURIComponent($("#hitb").val()));
            }

        });

        function checkchange(obj, id2) {
            if (treasurebox[$(obj).val().toString()].Piece == 1) {
                if ($(obj).val() == $(id2).val()) {
                    alert("您沒有兩個<" + treasurebox[$(obj).val().toString()].TreasureName + ">喔!!");
                    $(obj)[0].selectedIndex = 0;
                    return;
                }
                $(id2).children().each(function () {
                    if ($(this).val() == $(obj).val()) {
                        $(this).hide()
                    } else {
                        $(this).show()
                    }
                });
            } else {
                $(id2).children().each(function () {
                        $(this).show()
                });
                }
        }

        function check() {
            if($('#<%=TreasureDropDownListClientId%>').val() == 0 || $('#<%=TreasureDropDownList2ClientId%>').val() ==0 ){
                 alert("請確認選擇要兌換的寶物是否正確!!");
                 return false;
             }
             temp = "";
             if ($('#<%=TreasureDropDownListClientId%>').val() == $('#<%=TreasureDropDownList2ClientId%>').val()) {
                 temp = "您確認要使用兩個<" + $('#<%=TreasureDropDownListClientId%> option:selected').text() + ">交換隨機的一個寶物嗎?交換後不可以反悔喔!!"
             } else {
                temp = "您確認要使用一個<" + $('#<%=TreasureDropDownListClientId%> option:selected').text() + ">" + "與一個<" + $('#<%=TreasureDropDownList2ClientId%> option:selected').text() + ">交換隨機的一個寶物嗎?交換後不可以反悔喔!!";
             }
            if (confirm(temp))
             __doPostBack('ctl00$ContentPlaceHolder1$randomChangeTreasure', '');
        }

        function VoteLottery(objectId,voteValue){
            disableButton=  document.getElementById(objectId);
            if(confirm("兌換過後就不可反悔喔!!"))
            {
                disableButton.disabled = (disableButton.disabled == "" ? "disabled" : "");
                document.getElementById("hideVote").value=voteValue;
                document.forms[0].submit();
            }
        }
        function changeTreasure(id, icon,obj) {
            if (!changeflag)
                return;
            if ($('#<%=TreasureDropDownListClientId%>').val() == 0) {
                if (treasurebox[id].Piece == 1) {
                    if ($('#<%=TreasureDropDownList2ClientId%>').val() == id)
                        return;
                    if (treasurebox[id].Piece < 1)
                        return;
                }
                $('#<%=TreasureDropDownListClientId%>').children().each(function () {
                    if ($(this).val() == id) {
                        $(this).attr("selected", "selected");
                        $('#Img3').attr("src", "image/" + icon);
                        $('#Img1').fadeOut('slow', function () {

                        });
                        $('#Img3').fadeIn('slow', function () {
                            // 淡入動畫完成後會進來這
                        });
                        $(obj).addClass("trhdivchoose");
                    }
                });

            } else if ($('#<%=TreasureDropDownList2ClientId%>').val() == 0) {
                if (treasurebox[id].Piece == 1) {
                    if ($('#<%=TreasureDropDownListClientId%>').val() == id)
                        return;
                    if(treasurebox[id].Piece < 1)
                        return;
                }
                $('#<%=TreasureDropDownList2ClientId%>').children().each(function () {
                    if ($(this).val() == id) {
                        $(this).attr("selected", "selected");
                        $('#Img4').attr("src", "image/" + icon);
                        $('#Img2').fadeOut('slow', function () {

                        });
                        $('#Img4').fadeIn('slow', function () {
                            
                        });
                        $(obj).addClass("trhdivchoose");
                    }
                });
            }
        }
        function cancelThisT(id, se,id2) {
            if ($('#' + se).val() != 0) {
                $('#trediv' + $('#' + se).val()).removeClass("trhdivchoose");
                $('#' + se).val(0);
                $('#' + id).attr("src", "image/bean_q.gif");
                $('#' + id2).fadeOut('slow', function () {

                });
                $('#' + id).fadeIn('slow', function () {
                    // 淡入動畫完成後會進來這
                });
            }
        }
        function IwantChange(obj) {
            $(obj).hide();
            $('#changety').show();
            changeflag = true;
            tra =  $('[name=changetreasure]');
            $(tra).each(function () {
                $(this).addClass("pointer");
                $(this).mouseover(function () {
                    $(this).addClass("trhdivover");
                });
                $(this).mouseleave(function () {
                    $(this).removeClass("trhdivover");
                });
            });
        }
        function calcelChangeTreasure() {
            cancelThisT('Img1', '<%=TreasureDropDownListClientId%>', 'Img3')
            cancelThisT('Img2', '<%=TreasureDropDownList2ClientId%>', 'Img4'); 
            $('#iwantChangei').show();
            $('#changety').hide();
            
            changeflag = false;
            tra = document.getElementsByName("changetreasure");
            $(tra).each(function () {
                $(this).removeClass("pointer");
               
                $(this).removeClass("trhdivchoose");
                $(this).unbind();
            });

        }
        function doView(baseurl1, title) {
            // var baseurl = ""+baseurl1;


            GB_showCenter(title, baseurl1, /* optional */395, 490)
        }
    </script>
    <a title="網頁內容資料" class="Accesskey" accesskey="C" href="#" xmlns:hyweb="urn:gip-hyweb-com">
        :::</a>
    <div class="path" xmlns:hyweb="urn:gip-hyweb-com">
        目前位置：<a title="首頁" href="/mp.asp?mp=1">首頁</a>&gt;<a title="現代神農尋寶戰" href="#">現代神農尋寶戰</a></div>
        <asp:panel id="treasureall" runat="server" visible="true" >
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="treasuremenu"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="80" align="left" valign="bottom"><a href="01.aspx?activityid=3"><img src="image/mn_1.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="02.aspx?activityid=3"><img src="image/mn_2.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="03.aspx?activityid=3"><img src="image/mn_3.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="04_0.aspx?activityid=3"><img src="image/mn_4.gif" width="78" height="28" /></a></td>
    <td width="68" valign="bottom"><a href="Top.aspx?activityid=3"><img src="image/mn_5.gif" width="66" height="28" /></a></td>
    <td valign="bottom"><a href="#"><img src="image/mn_6h.gif" width="107" height="28" /></a></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td class="rank"><div class="bg" style="padding:15px 10px; text-align:center">
        <table width="100%" border="0" cellpadding="0" cellspacing="0" id="none">
          <tr>
            <td width="30" height="30" style="background:url(image/qa_br_tl.gif) no-repeat bottom">&nbsp;</td>
            <td align="center" style="background:url(image/qa_br_t.gif) repeat-x bottom"><img src="image/private_h1.gif" width="441" height="180"/></td>
            <td width="30" style="background:url(image/qa_br_tr.gif) no-repeat bottom">&nbsp;</td>
          </tr>
          <tr>
            <td style="background:url(image/qa_br_l.gif) repeat-y right">&nbsp;</td>
            <td style="background:#fff; padding:5px"><table width="100%" border="0" cellpadding="0" cellspacing="0" id="qa">
              <tr>
                <th width="20%" align="center">帳號</th>
                <td width="80%" rowspan="6">
                <asp:label ID="treasureTable" runat="server" text=""></asp:label>           
                <asp:panel id="acs" runat="server" visible="false">
                    <div style="font-size:14px; "><asp:label ID="canotchangeMessage" runat="server" text="您今天已經交換過寶物了喔!!"></asp:label></div>
                </asp:panel>
                 <asp:panel id="lotteryPane" runat="server" visible="false" >
                 <table width="100%" border="0" cellpadding="0" cellspacing="0" id="present" style="display:none;"> 
                    <tr> 
                      <td colspan="3" align="left"><h2>兌換專區</h2></td> 
                      </tr> 
                    <tr> 
                      <td  align="center"><img src="image/present_a2.jpg" width="80" height="80" /><p>幸運好禮一<br /> 
                          <span class="present_title">25L綠能小冰箱</span></p> 
                          <input type="button" name="submitLotteryA" id="submitLotteryA" onclick="VoteLottery('submitLotteryA','D');"  value="兌換一次抽獎機會" disabled="disabled" class="lotteryButton" > 
                        <p>共<asp:label ID="labelLotteryA" runat="server" text="0" CssClass="chance"></asp:label>次 </p></td> 
                      <td align="center"><img src="image/present_b2.jpg" width="80" height="80" /><p>幸運好禮二<br /> 
                          <span class="present_title">FUJIFILM 拍立得相機</span></p> 
                          <input type="button" name="submitLotteryB" id="submitLotteryB" onclick="VoteLottery('submitLotteryB','E');"  value="兌換一次抽獎機會" disabled="disabled" class="lotteryButton" > 
                        <p>共<asp:label ID="labelLotteryB" runat="server" text="0" CssClass="chance"></asp:label>次 </p></td> 
                      <td  align="center"><img src="image/present_c2.jpg" width="80" height="80" /><p>幸運好禮三<br /> 
                          <span class="present_title">柚花淨白系列</span></p> 
                          <input type="button" name="submitLotteryC" id="submitLotteryC" onclick="VoteLottery('submitLotteryC','F');" value="兌換一次抽獎機會" disabled="disabled" class="lotteryButton" > 
                        <p>共<asp:label ID="labelLotteryC" runat="server" text="0" CssClass="chance"></asp:label>次 </p></td> 
                    </tr> 
                  </table>
                   </asp:panel>
                </td>
                </tr>
              <tr align="center">
                <td class="private"><asp:label ID="loginid" runat="server" text=""></asp:label><br /></td>
              </tr>
              <tr align="center">
                <th>暱稱</th>
              </tr>
              <tr align="center">
                <td class="private"><asp:label ID="userName" runat="server" text=""></asp:label></td>
              </tr>
              <tr align="center">
                <th>可抽獎次數</th>
              </tr>
              <tr align="center">
                <td class="private"><asp:label ID="UsersTotalSuit" runat="server" text="0"></asp:label></td>
              </tr>
               <asp:panel id="changePane" runat="server" visible="false" >
              <tr align="center">
                <th>已兌換套數</th>
              </tr>
              <tr align="center">
                <td class="private"><asp:label ID="Lottery" runat="server" text="0"></asp:label></td>
              </tr>
              <tr align="center">
                <th>尚未兌換套數</th>
              </tr>
              <tr align="center">
                <td class="private"><asp:label ID="BeforeLottery" runat="server" text="0"></asp:label></td>
              </tr>
              </asp:panel>
              <tr id="changety" style=" display:none;">
              <th >交換專區</th>
              <td >
                <asp:panel id="randomPanel" runat="server" visible="true">
                     <div style="font-size:14px; text-align:left;padding-left:10px;"><div>
                  1.請點由上方寶物區選擇要交換的寶物 <br/>  
                  2.使用以下兩個寶物隨機換取一個寶物  <br/>  
                  3.點選下方寶物圖示可以取消選取的寶物<br/></div>
                  <table style=" border:medium nonel; width:200px;height:130px;position:relative; padding-left:5px;" align="left">
                  <tr style=" border:medium none;position:relative;">
                    <td  style=" border:medium none;position:relative;">
                        <div style=" padding: 0 ;position:relative;" >
                        <span style="position: absolute;" >
                        <span class="cti"><img   alt=""   src="/treasurehunt/2011/image/bean_q.gif"id="Img1" /></span>
                        <span class="cti"><img  alt="" onclick="cancelThisT('Img1','<%=TreasureDropDownListClientId%>','Img3');"  src="/treasurehunt/2011/image/bean_q.gif"id="Img3" /></span>
                        </span>
                        </div>
                    </td>
                     <td  style=" border:medium none;position:relative;">
                        <div style=" padding: 0 ;;position:relative;" >
                        <span style="position: absolute;" >
                        <span class="cti"><img alt=""   src="/treasurehunt/2011/image/bean_q.gif"id="Img2" /></span>
                        <span class="cti"><img alt=""  onclick="cancelThisT('Img2','<%=TreasureDropDownList2ClientId%>','Img4');"  src="/treasurehunt/2011/image/bean_q.gif"id="Img4" /></span>
                        </span>
                        </div>
                    </td>
                  </tr>
                  </table>
                  <div style=" display:none" >
                     <asp:DropDownList ID="TreasureDropDownList" autopostback="false" runat="server" visible="false"/>
                     <asp:DropDownList ID="TreasureDropDownList2" autopostback="false" runat="server" visible="false"/>
                  </div>
                  <table style=" border:medium none;text-align:right;height:130px;position:relative; vertical-align:bottom;">
                  <tr style=" border:medium none;">
                  <td width="100%"  style=" border:medium none;">
                  <div style=" border:medium none;text-align:left;"><br /><br /><br />
                    <asp:button runat="server" text="交換" id="randomChangeTreasure"  UseSubmitBehavior="false" />&nbsp;&nbsp;
                     <input type="button"  value="取消" id="calcelChangeTreasurei" onclick="calcelChangeTreasure()" />
                  </div>
                  </td>
                  </tr>
                  </table>
                  
                  <input type="hidden" id="hitb" value="<%=usersTreasureJson %>" />
                  </div>
                </asp:panel>
                </td>
              </tr>
            </table></td>
            <td style="background:url(image/qa_br_r.gif) repeat-y left">&nbsp;</td>
          </tr>
          <tr>
            <td height="30" style="background:url(image/qa_br_bl.gif) no-repeat top">&nbsp;</td>
            <td style="background:url(image/qa_br_b.gif) repeat-x top">&nbsp;</td>
            <td style="background:url(image/qa_br_br.gif) no-repeat top">&nbsp;</td>
          </tr>
        </table>
        <input type="hidden" value="" name="hideVote" id="hideVote" />
    </div></td>
  </tr>
</table>
</asp:panel>
 <asp:panel id="changebox" runat="server" visible="false" >
    <script type="text/javascript">
        $(document).ready(function (e) {
            doView('<%=changeTreasureBox%>', "");
        });
    </script>
 </asp:panel>
</asp:Content>

<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" MasterPageFile="Default.master" %>

<asp:Content ID="Content1" ContentPlaceHolderID="cp" runat="Server">
    <asp:HiddenField ID="useManagementMasterPage" runat="server" />
    <div id="content">
        <img src="images/COA_PlantGrowth_title-01.gif" width="243" height="42" />
        <p>
            本網站舉辦植物生長紀錄競賽活動，主要目的在於透過活動的進行，推廣家庭園藝，希望此活動能夠讓更多民眾體驗到植物種植的樂趣，也歡迎喜好家庭園藝的民眾，多多利用本網站資源，在此與同好們分享交流種植的經驗！
        </p>
        <img src="images/COA_PlantGrowth_title-02.gif" width="243" height="42" />
        <ul>
            <li>凡對於家庭園藝有興趣之個人，且具中華民國國籍，不限年齡、職業、性別，並加入成為本網站會員後，皆可報名參加。</li>
        </ul>
        <p class="center">
            <a href='<%= WebUtility.GetAppSetting("WWWUrl") %>/sp.asp?xdURL=coamember/member_Join.asp&mp=1'><img src="images/COA_PlantGrowth_join.gif" width="198" height="63" border="0" /></a></p>
        <img src="images/COA_PlantGrowth_title-04.gif" width="243" height="42" />
        <ul>
            <li>參賽者於競賽截止時間前，於活動網站上填寫作品介紹表，並上傳植物生長紀錄，即視為報名成功。 </li>
            <li>競賽結束後，符合競賽規則之作品將公開展示一週，由本站會員共同票選優秀作品。 </li>
        </ul>
        <img src="images/COA_PlantGrowth_title-03.gif" width="243" height="42" />
        <ul>
            <li>競賽進行時間：<span class="datestyle">98年7月30日中午12時起</span>，至<span class="datestyle">98年9月18日中午12時止</span>。 </li>
            <li>投票活動期間：<span class="datestyle">98年9月21日中午12時起</span>，至<span class="datestyle">98年9月28日中午12時止</span>。 </li>
            <li>相關抽獎活動將於投票活動後舉辦，得獎與中獎名單於投票活動截止後二週內公布於網站。</li>
        </ul>
        <img src="images/COA_PlantGrowth_title-07.gif" width="243" height="42" />
        <ol>
            <li>參賽者須於競賽活動截止時間<span class="datestyle">(98年9月18日中午12點)</span>之前，將植物生長紀錄全數上傳至網站，組成一份植物生長紀錄書。上傳流程說明<a href="images/UploadProcess.jpg" target="_blank">請按此</a>。</li>
            <li>合格的生長紀錄書至少要有四篇紀錄，每篇紀錄包括照片及文字部份，說明如下：
                <ul>
                    <li>照片：需提供實物拍攝、未經編輯修改之照片作為佐證。照片解析度建議為800*600以上。</li>
                    <li>文字：每篇紀錄需搭配種植文字說明，參賽者可介紹植物品種，種植情況(如溫度、土壤、水份、施肥、灌溉方式等等)，亦歡迎分享種植心得。生長紀錄上傳範例<a href="sample.htm" target="_blank">請點此</a>。</li>
                </ul>
            </li>
            <li>參賽者須選擇生長週期超過14天之植物，清楚、完整紀錄由種籽或阡插、萌芽、成長、成熟，開花、結果…等生長情形，合格之參賽作品，至少需呈現植物三個階段的變化。 </li>
            <li>參賽期間，每位參賽者至少需公開一筆生長紀錄。競賽活動結束時，所有生長紀錄將自動變為公開狀態，以利網站會員觀看票選。不符合競賽規定之參賽作品，將不列入票選名單中。 </li>
            <li>網站管理員擁有所有作品之管理及刪除權限。不符合競賽規定之參賽作品，將由管理員由票選名單中刪除。</li>
            <li>為防範有心人士運用活動網站功能散布不當之文字或圖片，凡發現參賽作品內容與活動無關，管理者可直接刪除相關紀錄，並對該帳號做出停權處分，取消參賽資格，不另行通知；參賽者不得異議。</li>
        </ol>
        <img src="images/COA_PlantGrowth_title-05.gif" width="243" height="42" />
        <p>
            本次競賽將由農業知識入口網之會員共同票選優秀作品。票選活動結束後，將依各作品得票數多寡，選出前3名得獎作品。獎勵如下：</p>
        <ul>
            <li>超人氣頭獎(1名)：Sony DSC-T900數位相機一台。</li>
            <li>超人氣貳獎(1名)：Apple ipod nano 16G一台。 </li>
            <li>超人氣參獎(1名)：Transend 500GB行動硬碟一個。 </li>
        </ul>
        <p>
            另外為鼓勵會員參與競賽，凡參加競賽活動，且未進入前三名之參賽者，可以參加參加獎抽獎活動，獎品如下：</p>
        <ul>
            <li>參加獎(10名)：8GB隨身碟一個。</li>
        </ul>
        <img src="images/COA_PlantGrowth_title-08.gif" width="243" height="42" />
        <p>
            為酬謝本網站會員協助進行優秀作品票選，凡網站會員曾參與投票，且未獲得競賽活動相關獎項者，可參加投票抽獎活動。投票規則如下：
        </p>
        <ol>
            <li>限本網站會員才有參與票選活動之資格。</li>
            <li>本網站會員每人可投三票，會員可以在投票活動期間，分次完成投票，不須一次投完三票。</li>
            <li>為求比賽公平起見，三票不得重複投給同一參賽者。</li>
            <li>為鼓勵會員投票，每投一票可獲得一次抽獎機會。 </li>
        </ol>
        <p>
            投票抽獎獎品如下：
        </p>
        <ul>
            <li>好幸運獎(1名)：Apple ipod shuffle 1G一台。 </li>
            <li>好口福獎(5名)：精美農產禮盒五份。 </li>
        </ul>
        <img src="images/COA_PlantGrowth_title-09.gif" width="243" height="42" />
        <ol>
            <li>為提供公平參賽權，參賽者1人最多以1件報名參賽。</li>
            <li>參賽作品應在本網站獨家發表，且不得再參加其他競賽。</li>
            <li>參賽作品應為原創，不得抄襲他人，參賽之照片亦應為參賽者根據其作品實物直接拍攝，不得做編輯或加工，以免影響票選結果。</li>
            <li>參賽者需自行保存參賽作品檔案，已上傳之作品，主辦單位擁有修改、重製、攝影、著作、各類型態媒體廣告宣傳、刊印、公開展示及商品化等權利均不另予通知及無償使用。</li>
            <li>如有任何因電腦、網路、電話、技術或不可歸責於主辦單位之事由，而使參加者登錄之資料有延遲、遺失、錯誤、無法辨識或毀損之情況，主辦單位不負任何法律責任，參加者亦不得因此異議。</li>
            <li>參賽者應尊重所有會員公開、公平之決定，對票選結果不得有異議。 </li>
            <li>公佈得獎名單後1週內，將以e-mail及電話方式通知得獎者，中獎民眾須於期限內依通知進行身份確認，逾期、或身份不符者，取消得獎資格。各獎項將於公佈名單後1個月內陸續寄送。</li>
            <li>得獎者須正確填寫領獎收據，並附上身分證影本以便進行核對領獎身分。得獎者個人資料僅供本次核對領獎身分用，不予退還，若贈品價值逾新台幣20,000元時，得獎者須負擔10%贈品稅。</li>
            <li>同一得獎者不拘獎項，限得獎一次，若有重複中獎情況，主辦單位得以取消其資格、採順位候補。</li>
            <li>得獎者須配合出席相關頒獎典禮及作品展示活動。 </li>
            <li>本次活動得獎贈品僅限郵寄台、澎、金、馬地區。</li>
            <li>凡報名參賽者，視為認同本辦法的各項內容及規定，本辦法如有未盡事宜，除依法律相關規定外，主辦單位保留修改之權利，若有未盡事宜，得隨時另行補充並公佈於活動網頁。 </li>
            <li>參賽者若有違反本競賽規則所列之規定，法律刑責將由參賽者自行負責，主辦單位得取消其參賽資格，若為得獎者則追回已頒發之相關獎品。若造成第三者之權益損失，參賽者得負完全法律責任，不得異議。 </li>
            <li>本活動如遇不可抗拒之因素而更改相關內容及辦法，將以本網站公告為依據；活動獎項亦以網站公告資料為準，如遇不可抗拒之因素，主辦單位保留更換其他等值獎項之權利。</li>
            <li>本活動所涉及的個人資料蒐集、運用以及關於個人資料的保密，將適用農委會網站的隱私權保護政策。</li>
        </ol>
    </div>

    <script type="text/javascript">
        $(document).ready(function() {
            //Tab Menu
            $("#tabInfo").addClass("btnDistrict now");
            $("#tabList").addClass("btnDistrict");
            $("#tabVote").addClass("btnDistrict");
            $("#tabFinal").addClass("btnDistrict");
        });
    </script>
	<script type="text/javascript">
var gaJsHost=(("https:" == document.location.protocol) ? "https://ssl." : "http://www.");
document.write(unescape("%3Cscript src='" + gaJsHost + "google-analytics.com/ga.js' type='text/javascript'%3E%3C/script%3E"));
</script>
<script type="text/javascript">
try {
		    var pageTracker = _gat._getTracker("UA-9195501-1");
		    pageTracker._trackPageview();
        }
        catch (err) {alert(err.description ); }   
</script>

</asp:Content>

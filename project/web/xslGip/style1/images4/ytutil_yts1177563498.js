function toogleWatcher(token, current_video_id)
{
	if(!self.sharing_active)
	{
		enableWatcherShare(token, current_video_id);
	}
	else
	{
		disableWatcherShare(token, current_video_id);
	}
}

function enableWatcherShare(token, current_video_id)
{
	fadein();
	var varg = "";
	if (current_video_id)
		varg = "&v="+current_video_id;
	getUrlXMLResponse("/watcher?action_start_share=1" + varg + "&t="+ token, showEnabledWatcher);
}

function disableWatcherShare(token, current_video_id)
{
	var varg = "";
	if (current_video_id)
		varg = "&v="+current_video_id;
	getUrlXMLResponse("/watcher?action_stop_share=1"+varg+"&t="+token, showDisabledWatcher);
}

function showEnabledWatcher()
{

	self.sharing_active = true;
	//hideDiv('sharingSettingBox');
	var img = document.getElementById("sharingImg")
	img.src = '/img/icn_green_on_12x12.gif';
	img.title = "Active Sharing On";
	//document.getElementById("WatcherLink").innerHTML = 'Sharing: <b>on</b>';
	fadeout();
	if (document.getElementById("activesharing_start_button")) {
		hideDiv('activesharing_start_button');
		showDiv('activesharing_stop_button');
	}

}

function showDisabledWatcher()
{
	self.sharing_active = false;
	//hideDiv('sharingSettingBox');
	var img = document.getElementById("sharingImg")
	img.src = '/img/icn_red_off_12x12.gif';
	img.title = "Active Sharing Off";
	//document.getElementById("WatcherLink").innerHTML = 'Sharing: <b>off</b>';
        if (document.getElementById("activesharing_start_button")) {
                showDiv('activesharing_start_button');
                hideDiv('activesharing_stop_button');
        }
}


function makearray(n) {
	this.length = n;
    	for(var i = 1; i <= n; i++)
        this[i] = 0;
    return this;
}
hexa = new makearray(16);
for(var i = 0; i < 10; i++)
    hexa[i] = i;
hexa[10]="a"; hexa[11]="b"; hexa[12]="c";
hexa[13]="d"; hexa[14]="e"; hexa[15]="f";
function hex(i) {
    if (i < 0)
        return "00";
    else if (i > 255)
        return "ff";
    else
        return "" + hexa[Math.floor(i/16)] + hexa[i%16];
}

function setbgColor(r, g, b) {
    var hr = hex(r); var hg = hex(g); var hb = hex(b);
    document.getElementById('shareSpan').style.backgroundColor = "#" +hr+hg+hb;
}

function fade(sr, sg, sb, er, eg, eb, step) {
    for(var i = 0; i <= step; i++) {
       setbgColor(
        Math.floor(sr * ((step-i)/step) + er * (i/step)),
        Math.floor(sg * ((step-i)/step) + eg * (i/step)),
        Math.floor(sb * ((step-i)/step) + eb * (i/step)));
    }
}


function fadein() {
	fade(255,255,255, 102,204,0, 825);
}

function fadeout() {
	setTimeout("fade(102,204,0, 255,255,255,2025)",800)
}
// -->

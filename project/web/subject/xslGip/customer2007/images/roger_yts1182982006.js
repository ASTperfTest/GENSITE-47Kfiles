var req;
if( window != window.top ) { 
	var ref = document.referrer;
	var sel = window.location.href;
	req = getXmlHttpRequest();
	var data = 'location=' + encodeURIComponent(ref);
	data += '&self=' + encodeURIComponent(sel);
	req.onreadystatechange = processReqChange;
	req.open("POST", '/roger_rabbit', true);
	req.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
	req.send(data);
}
function processReqChange() {
    if (req.readyState == 4) {
        if (req.status == 200) {
		if (req.responseText == 'block') {
			window.top.location.href = '/';
		}
        } 
    }
}

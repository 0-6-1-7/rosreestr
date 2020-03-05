var oldCaptchaSrc = "";

let timerId = setInterval(() => getCaptcha(), 500);

function httpGetAsync(imgData, eleId) {
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.open("GET", "http://localhost:8111/cgi-bin/captcha.py?imgData='" + imgData + "'", true);
	xmlHttp.onload = function() { eleId.value = this.responseText; };
	//xmlHttp.onerror = function() { eleId.value = "-----"; };
	xmlHttp.onerror = function() { eleId.value = ""; };
    xmlHttp.send(null);
}

function getBase64Image(img) {
    var canvas = document.createElement("canvas");
    canvas.width = img.width;
    canvas.height = img.height;
    var ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0);
    var dataURL = canvas.toDataURL("image/png");
    return dataURL.replace(/^data:image\/(png|jpg);base64,/, "");
	canvas = null
}

function base64_url_encode(url) {
 return url.replace(/\+/g, ".").replace(/\//g, "_").replace(/\=/g, "-");
}

function getCaptcha() {
	if (document.title.length < 1) { return; }
	if (document.URL.indexOf("https://rosreestr.ru/wps/portal/online_request") >= 0 ||
		document.URL.indexOf("https://rosreestr.ru/wps/portal/p/cc_ib_portal_services/online_request/") >= 0) {
		var captcha = document.getElementById("captchaImage2");
		var captchaField = document.getElementsByName("captchaText")[0];
	}
	else if (document.URL.indexOf("https://rosreestr.ru/wps/portal/p/cc_present/ir_egrn") >= 0){
		var ibmMainContainer = document.getElementsByName("ibmMainContainer")[0];
		if (ibmMainContainer.getElementById("gwt-uid-3") == null) { return; }
		var captcha = ibmMainContainer.querySelector("img");
		var captchaField = ibmMainContainer.querySelector("input");
		ibmMainContainer.getElementById("gwt-uid-3").click();
		ibmMainContainer.querySelector("div.v-button").focus();
	} 
	else {return;}
	try {
		if(!captcha.complete) { return; }
		if (oldCaptchaSrc != captcha.src) {
			oldCaptchaSrc = captcha.src;
			solvedCaptcha = httpGetAsync(base64_url_encode(getBase64Image(captcha)), captchaField);
		}
	} catch { return; }
}


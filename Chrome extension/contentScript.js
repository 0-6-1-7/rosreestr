var oldCaptchaSrc = "";

function httpGetAsync(imgData, eleId) {
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.open("GET", "http://localhost:8111/cgi-bin/captcha.py?imgData='" + imgData + "'", true);
	xmlHttp.onload = function() { 
		eleId.focus(); 
		setTimeout(() => { eleId.value = this.responseText; }, 100);
	}
	xmlHttp.onerror = function() { eleId.value = "-----"; };
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
	canvas = null;
}

function base64_url_encode(url) {
 return url.replace(/\+/g, ".").replace(/\//g, "_").replace(/\=/g, "-");
}

function getCaptcha() {
	if (document.title.length < 1) { return; } //сайт не работает
	captcha = null;
	var container = document.querySelector("div.content-right");
	var imgs = container.querySelectorAll("img");
	for (let img of imgs) { if (img.width == 180 && img.height == 50) { captcha = img; break; } }
	if (captcha == null) { return; }
	
	if (document.URL.indexOf("online_request") >= 0 ||
		document.URL.indexOf("cc_vizualisation") >= 0) {
		var captchaField = captcha.closest("tbody").querySelector("input");
		}
	else if (document.URL.indexOf("cc_present/ir_egrn") >= 0 ||
		     document.URL.indexOf("cc_present/EGRN_") >= 0) {
		var captchaField = captcha.closest("div.v-horizontallayout").querySelector("input");
	} 
	else {return;}
	if(!captcha.complete) { return; }
	if (oldCaptchaSrc != captcha.src) {
		oldCaptchaSrc = captcha.src;
		solvedCaptcha = httpGetAsync(base64_url_encode(getBase64Image(captcha)), captchaField);
	}
}

let timerId = setInterval(() => getCaptcha(), 500);
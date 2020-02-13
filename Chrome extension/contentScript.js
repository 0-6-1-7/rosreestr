var oldCaptchaSrc = "";

function httpGetAsync(imgData, eleId) {
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.open("GET", "http://localhost:8111/cgi-bin/captcha.py?imgData='" + imgData + "'", true);
	xmlHttp.onload = function() { eleId.value = this.responseText; };
	xmlHttp.onerror = function() { eleId.value = '-----'; };
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
 return url.replace(/\+/g, '.').replace(/\//g, '_').replace(/\=/g, '-');
}

function getCaptcha() {
	var captcha = document.getElementById("captchaImage2");
	if(!captcha.complete) { console.log('image not complete'); return; }
	if (oldCaptchaSrc != captcha.src) {
		oldCaptchaSrc = captcha.src;
		solvedCaptcha = httpGetAsync(base64_url_encode(getBase64Image(captcha)), document.getElementsByName("captchaText")[0]);
	}
}

let timerId = setInterval(() => getCaptcha(), 100);

chrome.runtime.onMessage.addListener(Listener);

try { if (r_enc.style.display == "none") { sw_r_enc.click(); } }  catch {}
try { if (s_notes.style.display == "none") { sw_s_notes.click(); } } catch {}


function Listener(request, sender, sendResponse) {
		if (request.storedata == "storedata")
			payload(function() { sendResponse("goodbye");
			return true; });
}
	
function payload(callback){
	var KN;
	const re = /\d{2}:\d{2}:\d{1,}[:\d{1,}]{1,}/gm;

	try { if (r_enc.style.display == "none") { sw_r_enc.click(); } }  catch {}
	try { if (s_notes.style.display == "none") { sw_s_notes.click(); } } catch {}
	
	var Content = document.querySelector("div.portlet-body").innerText + "\n";

	if (Content.indexOf("Кадастровый номер:") !== -1) { 
		KN = re.exec(Content)[0];
		chrome.storage.local.set({ [KN]: Content }, function() {});

		// navigator.permissions.query({ name: 'clipboard-write' }).then(permissionStatus => {
			// console.log(permissionStatus.state);
			// navigator.clipboard.writeText(s);
	};
	callback();
}

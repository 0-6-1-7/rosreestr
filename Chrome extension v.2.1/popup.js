document.addEventListener('DOMContentLoaded', function () {
  document.getElementById('ClearStorage')
    .addEventListener('click', ClearStorage);
});

document.addEventListener('DOMContentLoaded', function () {
  document.getElementById('SaveContent')
    .addEventListener('click', SaveContentHandler);
});

function ClearStorage() {
	Content.value=''; Counter.innerText = "0"; chrome.storage.local.clear();
}

function GetStorage() {
	chrome.storage.local.get(null, function(items) { 
		Content.value = JSON.stringify(items); 
		try {
			Counter.innerText = Object.keys(items).length;
		} catch {}
		});
}

function SaveContentHandler() {
	chrome.tabs.query( {active: true, currentWindow: true}, 
		function(tabs) { chrome.tabs.sendMessage( tabs[0].id, {storedata: "storedata"}, function() { GetStorage(); } );
		}
	);
}
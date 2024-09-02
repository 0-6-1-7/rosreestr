const rr = 'https://lk.rosreestr.ru/eservices/real-estate-objects-online';

chrome.runtime.onInstalled.addListener(() => {
	chrome.action.setBadgeText({ text: 'Выкл', });
	chrome.storage.local.set({ state: 'Выкл' }).then(() => {});
	
});


chrome.action.onClicked.addListener(async (tab) => {
	if (tab.url.startsWith(rr)) {
		const prevState = await chrome.action.getBadgeText({ tabId: tab.id });
		let nextState = '';
		switch (prevState) {
			case 'Выкл': 
				nextState = '34';
				break;
			case '34':
				nextState = 'i';
				break;
			case 'i':
				nextState = 'Выкл';
				break;
		};

		await chrome.action.setBadgeText({
			tabId: tab.id,
			text: nextState,
		});
		
		const response = await chrome.tabs.sendMessage(tab.id, {state: nextState});

		chrome.storage.local.set({ state: nextState }).then(() => {});
			
		if (nextState === '34') { //inject CSS just once
			await chrome.scripting.insertCSS({
				files: ['styles/focus-mode.css'],
				target: { tabId: tab.id },
			});
		} else if (nextState === 'Выкл') {
			try {
				await chrome.scripting.removeCSS({
					files: ['styles/focus-mode.css'],
					target: { tabId: tab.id },
				});
			} catch (err) {
				console.error(`failed to remove CSS: ${err}`);
			};
		};
	};
});

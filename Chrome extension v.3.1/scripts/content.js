let intervalWaitModal = null;
let intervalWaitErrors = null;
let workMode = 'Выкл'; //'' = Выкл, i, 34

chrome.runtime.onMessage.addListener(
	function(request, sender, sendResponse) {
		workMode = request.state;
		waitModal(true); //force - если переключаем режим рабоиы из сведений, принудительно копируется КН или инфо
	}
);

function clearInfo(i) {
	return String(i).replace('\nДЕЙСТВИЯ\n', '\n');
}

async function writeClipboardText(element1, element2) {
//element1 - кадастровый номер (для раскрашивания), element2 - элемент, из которго будет взят innerText для буфера
	try {
		const textToCopy = String(element2.innerText).replace(/\n$/, '');
		await navigator.clipboard.writeText(textToCopy);
		element1.setAttribute('kn-status', 'copied');
	} catch (error) {
		//console.error(`Exception. ${error.name}: ${error.message}`);
		element1.setAttribute('kn-status', 'error');
	}
}

function waitModal (force) {
	if (workMode === 'Выкл') { return };
	const m = document.querySelector('button.rros-ui-lib-modal__close-btn');
	if (m && m.checkVisibility()) {
		try {
			let poe = false; //наличие раздела 'Сведения о правах и ограничениях (обременениях)'
			let soe = null; //строка 'Статус объекта'
			let soeText = ''; //значение в строке 'Статус объекта' 
			let kne = null; //строка 'Кадастровый номер' в разделе 'Общая информация'
			let kneColor = ''; // цвет кадастрового номера
			let etc = null; //элемент, из которго будет взят innerText для буфера (кадстровый номер или вся карточка)
			elements = document.querySelectorAll('li.build-card-wrapper__info__ul__subinfo');
			for (const e of elements) {
				const h = e.querySelector('span.build-card-wrapper__info__ul__subinfo__name');
				switch (h.innerText) {
					case 'Кадастровый номер':
						if (!kne) {
							kne = e;
						};
						break;
					case 'Статус объекта':
						soe = e;
						soeText = e.querySelector('div.build-card-wrapper__info__ul__subinfo__options__item__line').innerText;
						break;
					case 'Вид, номер и дата государственной регистрации права':
						poe = true;
						break;
					default:
				};
			};
			switch (soeText) {
				case 'Актуально':
					if (poe) {
						soe.setAttribute('so-type', 'active')
						kneColor = kne.style.backgroundColor;
						knStatus = kne.getAttribute('kn-status');
						if (knStatus != 'copied' || force) {
							if (window.isSecureContext) {
								switch (workMode) {
									case '34':
										etc = kne.querySelector('div.build-card-wrapper__info__ul__subinfo__options__item__line');
										break;
									case 'i':
										etc = document.querySelector('div.build-card-wrapper.card')
										break;
								};
								try {
									writeClipboardText(kne, etc);
								} catch (error) {
									console.error(`Exception. ${error.name}: ${error.message}`);
								};
							};
						};
					} else {
						soe.setAttribute('so-type', 'inactive')
					};
					break;
				case 'Погашено':
					soe.setAttribute('so-type', 'closed')

					break;
			}
		} catch (error) {
			console.error(`Exception in waitModal. ${error.name}: ${error.message}`);
		};
	};
}


function waitErrors () {
	if (workMode === 'Выкл') { return };
	const errors = document.querySelectorAll('div.rros-ui-lib-error.rros-ui-lib-error-ERROR');
	if (errors.length > 0) {
		for (const e of errors) {
			console.log(`Сообщение сайта: ${e.innerText}`);
			try {
				const b = e.querySelector('button.rros-ui-lib-button.rros-ui-lib-button--link');
				b.click()
			} catch (error) {
				console.error(`Exception in waitErrors. ${error.name}: ${error.message}`);
			};
		};
	};
}

function main() {
	intervalWaitModal = window.setInterval(waitModal, 300);
	intervalWaitErrors = window.setInterval(waitErrors, 5000);
}

main();

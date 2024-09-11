let intervalWaitModal = null;
let intervalWaitErrors = null;
let workMode = 'Выкл'; //'' = Выкл, i, П, 34

chrome.runtime.onMessage.addListener(
	function(request, sender, sendResponse) {
		workMode = request.state;
		switch (workMode) {
			case 'Выкл':
				removeButtonPaste();
				break;
			case '34':
			case ' П ':
			case ' i ':
				let button = document.getElementById('my-button-paste');
				if (!button) { insertButtonPaste() };
				break;
		};
			
		waitModal(true); //force - если переключаем режим работы из сведений, принудительно копируется КН или инфо
	}
);

function insertButtonPaste () {
	let element = document.querySelector('input.realestateobjects-wrapper__option_input');
	if (!element) { return };
	const buttonSVG = '<svg xmlns="http://www.w3.org/2000/svg" height="24px" viewBox="0 -960 960 960" width="24px" fill="#5f6368"><path d="M200-120q-33 0-56.5-23.5T120-200v-560q0-33 23.5-56.5T200-840h167q11-35 43-57.5t70-22.5q40 0 71.5 22.5T594-840h166q33 0 56.5 23.5T840-760v560q0 33-23.5 56.5T760-120H200Zm0-80h560v-560h-80v120H280v-120h-80v560Zm280-560q17 0 28.5-11.5T520-800q0-17-11.5-28.5T480-840q-17 0-28.5 11.5T440-800q0 17 11.5 28.5T480-760Z"/></svg>';
	let button = document.createElement('span');
	button.id = 'my-button-paste';
	button.setAttribute('title', 'Вставить в поле текст из буфера обмена');
	button.classList.add('my-button', 'my-button-adj-right');
	element.insertAdjacentElement('afterend', button);
	button.insertAdjacentHTML('afterbegin', buttonSVG);
	button.addEventListener('click', clickButtonPaste);	
}

function removeButtonPaste() {
	const button = document.getElementById('my-button-paste');
	button.remove();
}

function clickButtonPaste(event) {
	const input = document.getElementById('query');
//	const button = document.getElementById('realestateobjects-search');
//	console.log(button);
	navigator.clipboard.readText().then((clipText) => {
		const regex = /^ *\d{2} *: *\d{2} *: *\d{1,7} *: *\d+ *$/;
		let text = String(clipText);
		if (text.match(regex)) { text = text.replace(/ /g, '') };
		input.focus();
		document.execCommand('selectAll');
		document.execCommand('insertText', false, text);
//			console.log(clipText);
//			button.click();
	});
}

function insertButtonsCopy() {
	const buttonSVG = '<svg xmlns="http://www.w3.org/2000/svg" height="24px" viewBox="0 -960 960 960" width="24px" fill="#5f6368"><path d="M360-240q-33 0-56.5-23.5T280-320v-480q0-33 23.5-56.5T360-880h360q33 0 56.5 23.5T800-800v480q0 33-23.5 56.5T720-240H360Zm0-80h360v-480H360v480ZM200-80q-33 0-56.5-23.5T120-160v-560h80v560h440v80H200Zm160-240v-480 480Z"/></svg>';

	let element = null;
	let button = null;
	
	//Кадастровый номер
	button = document.getElementById('my-button-copy-1');
	if (!button) {
		element = getElementKN();
		button = document.createElement('span');
		button.id = 'my-button-copy-1';
		button.setAttribute('title', 'Скопировать кадастровый номер');
		button.classList.add('my-button', 'my-button-adj-left');
		element.insertAdjacentElement('afterbegin', button);
		button.insertAdjacentHTML('afterbegin', buttonSVG);
		button.addEventListener('click', clickButtonCopy1);
	};

	//Сведения о правах
	button = document.getElementById('my-button-copy-2');
	if (!button) {
		element = getElementPO();
		button = document.createElement('span');
		button.id = 'my-button-copy-2';
		button.setAttribute('title', 'Скопировать сведения о правах');
		button.classList.add('my-button', 'my-button-adj-left');
		element.insertAdjacentElement('afterbegin', button);
		button.insertAdjacentHTML('afterbegin', buttonSVG);
		button.addEventListener('click', clickButtonCopy2);
	};

	//Сведения об объекте
	button = document.getElementById('my-button-copy-3');
	if (!button) {
		element = document.querySelector('h1');
	//	if (!e3) { return };
		button = document.createElement('span');
		button.id = 'my-button-copy-3';
		button.setAttribute('title', 'Скопировать полностью все данные');
		button.classList.add('my-button', 'my-button-adj-left');
		element.insertAdjacentElement('afterbegin', button);
		button.insertAdjacentHTML('afterbegin', buttonSVG);
		button.addEventListener('click', clickButtonCopy3);	
	};
}

function clickButtonCopy1(e) {
	const element = getElementKN();
	if (element) {
		const etc = element.querySelector('div.build-card-wrapper__info__ul__subinfo__options__item__line');
		const textToCopy = clearInfo1(etc.innerText);
		writeClipboardText(null, textToCopy);
	};
}

function clickButtonCopy2(e) {
	const element = getElementPO();
	if (element) {
		const etc = element.querySelector('div.build-card-wrapper__info__ul__subinfo__options');
		const textToCopy = clearInfo2(etc.innerText);
		writeClipboardText(null, textToCopy);
	};
}

function getElementKN() {
	const elements = document.querySelectorAll('li.build-card-wrapper__info__ul__subinfo');
	let kne = null; //элемент "Кадастровый номер"
	for (const element of elements) {
		const h = element.querySelector('span.build-card-wrapper__info__ul__subinfo__name');
		if (h.innerText === 'Кадастровый номер') {
			kne = element;
			break;
		};
	};
	return kne;
}

function getElementPO() {
	const elements = document.querySelectorAll('li.build-card-wrapper__info__ul__subinfo');
	let poe = null; //элемент "Вид, номер и дата государственной регистрации права"
	for (const element of elements) {
		const h = element.querySelector('span.build-card-wrapper__info__ul__subinfo__name');
		if (h.innerText === 'Вид, номер и дата государственной регистрации права') {
			poe = element;
			break;
		};
	};
	return poe;
}

function clickButtonCopy3(e) {
	const etc = document.querySelector('div.build-card-wrapper.card');
	const textToCopy = clearInfo3(etc.innerText);
	writeClipboardText(null, textToCopy);
}
function removeButtonsCopy() {
	let button = null;
	
	button = document.getElementById('my-button-copy-1');
	if (button) { button.remove() };
	
	button = document.getElementById('my-button-copy-2');
	if (button) { button.remove() };
	
	button = document.getElementById('my-button-copy-3');
	if (button) { button.remove() };
}

//34
function clearInfo1(text) {
	let t = String(text).replace(/\n$/, '');
	return t;
}

//П
function clearInfo2(text) {
	let t = String(text);
	t = t.replace(/(собственность)\n/gi, '$1 ');
	t = t.replace(/(управление)\n/gi, '$1 ');
	t = t.replace(/\n(от)/gi, ' $1');
	return t;
}

//i
function clearInfo3(text) {
	let t = String(text);
	t = t.replace('\nДЕЙСТВИЯ\n', '\n');

	t = t.replace(/(\nОбщая информация\n)/gi, '\n$1');
	t = t.replace(/(\nХарактеристики объекта\n)/gi, '\n$1');
	t = t.replace(/(\nСведения о кадастровой стоимости\n)/gi, '\n$1');
	t = t.replace(/(\nРанее присвоенные номера\n)/gi, '\n$1');
	t = t.replace(/(\nСведения о правах и ограничениях \(обременениях\)\n)/gi, '\n$1');

	t = t.replace(/(Дата обновления информации:)\n/gi, '$1 ');
	t = t.replace(/(Вид объекта недвижимости)\n/gi, '$1 ');
	t = t.replace(/(Статус объекта)\n/gi, '$1 ');
	t = t.replace(/(Кадастровый номер)\n/gi, '$1 ');
	t = t.replace(/(Дата присвоения кадастрового номера)\n/gi, '$1 ');
	t = t.replace(/(Форма собственности)\n/gi, '$1 ');
	t = t.replace(/(Адрес \(местоположение\))\n/gi, '$1 ');
	t = t.replace(/(Площадь, .+)\n/gi, '$1 ');
	t = t.replace(/(Назначение)\n/gi, '$1 ');
	t = t.replace(/(Этаж)\n/gi, '$1 ');
	t = t.replace(/(Количество этажей)\n/gi, '$1 ');
	t = t.replace(/(Количество подземных этажей)\n/gi, '$1 ');
	t = t.replace(/(Материал наружных стен)\n/gi, '$1 ');
	t = t.replace(/(Год завершения строительства)\n/gi, '$1 ');
	t = t.replace(/(Кадастровая стоимость .*)\n/gi, '$1 ');
	t = t.replace(/(Дата определения)\n/gi, '$1 ');
	t = t.replace(/(Дата внесения)\n/gi, '$1 ');
	t = t.replace(/(Инвентарный номер)\n/gi, '$1 ');

	t = t.replace(/(собственность)\n/gi, '$1 ');
	t = t.replace(/(управление)\n/gi, '$1 ');
	t = t.replace(/\n(от)/gi, ' $1');
//	t = t.replace(/()\n/gi, '$1 ');
	return t;
}


async function writeClipboardText(element, textToCopy) {
//element - кадастровый номер (для раскрашивания), textToCopy - текст для буфера обмена
	try {
		await navigator.clipboard.writeText(textToCopy);
		if (element) { element.setAttribute('kn-status', 'copied') };
	} catch (err) {
		if (element) { element.setAttribute('kn-status', 'error') };
	};
}

function waitModal (force) {
	if (workMode === 'Выкл') { 
		removeButtonsCopy();
		return;
	};
	const m = document.querySelector('button.rros-ui-lib-modal__close-btn');
	if (m && m.checkVisibility()) {
//		try {
			let poe = null; //раздел 'Сведения о правах и ограничениях (обременениях)'
			let soe = null; //строка 'Статус объекта'
			let soeText = ''; //значение в строке 'Статус объекта' 
			let kne = null; //строка 'Кадастровый номер' в разделе 'Общая информация'
			let kneColor = ''; // цвет кадастрового номера
			let etc = null; //элемент, из которго будет взят innerText для буфера (кадстровый номер или вся карточка)
			let textToCopy = ''; //текст для буфера омена из элемента etc
			const elements = document.querySelectorAll('li.build-card-wrapper__info__ul__subinfo');
			for (const element of elements) {
				const h = element.querySelector('span.build-card-wrapper__info__ul__subinfo__name');
				switch (h.innerText) {
					case 'Кадастровый номер':
						if (!kne) { kne = element };
						break;
					case 'Статус объекта':
						soe = element;
						soeText = element.querySelector('div.build-card-wrapper__info__ul__subinfo__options__item__line').innerText;
						break;
					case 'Вид, номер и дата государственной регистрации права':
						poe = element;
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
										textToCopy = clearInfo1(etc.innerText);
										break;
									case ' П ':
										etc = poe.querySelector('div.build-card-wrapper__info__ul__subinfo__options');
										textToCopy = clearInfo2(etc.innerText);
										break;
									case ' i ':
										etc = document.querySelector('div.build-card-wrapper.card');
										textToCopy = clearInfo3(etc.innerText);
										break;
								};
								writeClipboardText(kne, textToCopy);
							};
						} else {
							insertButtonsCopy();
						};
					} else {
						soe.setAttribute('so-type', 'inactive')
					};
					break;
				case 'Погашено':
					soe.setAttribute('so-type', 'closed')
					break;
			};
//		} catch (err) {
//			console.error(`Exception in waitModal: ${err.name}: ${err.message}`);
//		};
	} else {
		removeButtonsCopy();
	};
}


function waitErrors () {
	if (workMode === 'Выкл') { return };
	const errors = document.querySelectorAll('div.rros-ui-lib-error.rros-ui-lib-error-ERROR');
	if (errors.length > 0) {
		for (const error of errors) {
			console.log(`Сообщение сайта: ${error.innerText}`);
			try {
				const b = error.querySelector('button.rros-ui-lib-button.rros-ui-lib-button--link');
				b.click()
			} catch (err) {
				console.error(`Exception in waitErrors. ${err.name}: ${err.message}`);
			};
		};
	};
}

function main() {
	intervalWaitModal = window.setInterval(waitModal, 300);
	intervalWaitErrors = window.setInterval(waitErrors, 5000);
}

main();

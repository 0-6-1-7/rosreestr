async function writeClipboardText(text) {
	try {
		await navigator.clipboard.writeText(text);
	} catch (error) {
		console.error(`Exception. ${error.name}: ${error.message}`);
	}
}

function waitDOM () {
	const b = document.querySelector('#realestateobjects-search');
	if (b) {
		//console.log('waitDOM');
		document.querySelector('#headerNavLinks').style.display = 'none';
		let e = document.querySelector('div.rros-ui-lib-title.rros-ui-lib-title--1');
			e.style.marginTop = 0;
			e.style.marginBottom = 0;
			e.style.fontSize = '18px'
		e = document.querySelector('div.realestateobjects-wrapper.card');
				e.style.marginBottom = 0;
				e.style.paddingTop = 0;
				e.style.paddingBottom = 0;
		for (const e of document.querySelectorAll('div.realestateobjects-wrapper__option')) {
				e.style.marginBottom = 0;
				e.style.paddingTop = 0;
				e.style.paddingBottom = 0;
		};
		for (const e of document.querySelectorAll('label.rros-ui-lib-input-wrapper')) {
				e.style.marginBottom = 0;
				e.style.paddingTop = 0;
				e.style.paddingBottom = 0;
		};
	};
		setTimeout(waitDOM, 300); // 300 milliseconds
}

function waitResults () {
	const b = document.querySelector('div.row.realestateobjects-wrapper__results');
	if (b) {
		//console.log('waitResults');
		b.style = 'margin-top: 0px;';
		try {
			document.querySelector('div.rros-ui-lib-table__header').style = 'display: none;';
			const cells = document.querySelectorAll('div.rros-ui-lib-table__cell');
			for(const e of cells) {
				e.style.paddingTop = '10px';
				e.style.paddingBottom = '10px';
			};
		} catch (error) {
			console.error(`Exception. ${error.name}: ${error.message}`);
		};
	};
	setTimeout(waitResults, 300);
}

function waitModal () {
	const m = document.querySelector('button.rros-ui-lib-modal__close-btn');
	if (m && m.checkVisibility()) {
		//console.log('waitModal');
		try {
			document.querySelector('div.realestate-object-modal').style.marginTop = 0;
			document.querySelector('h1.realestate-object-modal__h1').style.marginBottom = 0;
			for (const e of document.querySelectorAll('h3')) {
				e.style.marginBottom = 0;
			};
			for (const e of document.querySelectorAll('div.build-card-wrapper__info')) {
				e.style.marginBottom = 0;
			};
			for (const e of document.querySelectorAll('ul.build-card-wrapper__info__ul')) {
				e.style.marginBottom = 0;
			};
			for (const e of document.querySelectorAll('li.build-card-wrapper__info__ul__subinfo')) {
				e.style.marginBottom = 0;
			};
			let poe = false;
			let soe = false;
			let kne;
			let kneColor;
			elements = document.querySelectorAll('li.build-card-wrapper__info__ul__subinfo');
			for (const e of elements) {
				const h = e.querySelector('span.build-card-wrapper__info__ul__subinfo__name');
				switch (h.innerText) {
					case 'Кадастровый номер':
						kne = e;
						break;
					case 'Статус объекта':
						soe = e.querySelector('div.build-card-wrapper__info__ul__subinfo__options__item__line').innerText == 'Актуально';
						break;
					case 'Вид, номер и дата государственной регистрации права':
						poe = true;
						break;
					default:
				};
			};
			if (soe && poe) {
				kneColor = kne.style.backgroundColor
				if (kneColor == '') {
					if (window.isSecureContext) {
						const kn = kne.querySelector('div.build-card-wrapper__info__ul__subinfo__options__item__line').innerText;
						try {
							writeClipboardText(kn);
						 } catch (error) {
							console.error(`Exception. ${error.name}: ${error.message}`);
						 };
						 kne.style.backgroundColor = '#55bb55'
					};
				};
			};
		} catch (error) {
			console.error(`Exception. ${error.name}: ${error.message}`);
		};
	};
	setTimeout(waitModal, 300);
}

function waitErrors () {
	const errors = document.querySelectorAll('div.rros-ui-lib-error.rros-ui-lib-error-ERROR');
	if (errors.length > 0) {
		for (const e of errors) {
			console.log(`Сообщение сайта: ${e.innerText}`);
			try {
				const b = e.querySelector('button.rros-ui-lib-button.rros-ui-lib-button--link');
				b.click()
			} catch (error) {
				console.error(`Exception. ${error.name}: ${error.message}`);
			};
		};
	};
	setTimeout(waitErrors, 5000);
}


//main
waitDOM();
waitResults();
waitModal();
waitErrors();



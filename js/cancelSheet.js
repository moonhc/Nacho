function createExcel(cancelData) {
	return createWorkbook().then(
		function (workbook) {
			createCancelSheets(workbook, cancelData);
		});
}

function createCancelSheets(workbook, cancelData) {
	let refinedData = refineCancelData(cancelData);

	for (let type in refinedData) {
		let sheetName = `${type}(취소)`;
		const sheet = workbook.addSheet(sheetName);
		writeToSheet(sheet, refinedData[type]);
	}
}

function writeToSheet(sheet, data) {
	return workbook;
}

function downloadExcel(workbook, fileName) {
	// Error handling for file name
	if (fileName.indexOf('.') == -1) {
		alert(`File name has no extension`);
		return;
	}
	let splitFileName = fileName.split('.');
	let extension = splitFileName[splitFileName.length-1];
	if (extension != 'xlsx' && extension != 'xls') {
		alert(`File extension (${extension}) is not supported`);
		return;
	}

	// Download file with given file name
	workbook.outputAsync().then(
		function (blob) {
			let url = window.URL.createObjectURL(blob);
			let a = document.createElement('a');
			document.body.appendChild(a);
			a.href = url;
			a.download = fileName;
			a.click();
			window.URL.revokeObjecURL(url);
			document.body.removeChild(a);
		}).catch(
		function (err) {
			alert(err.message || err);
		});
}

function findSheetByName(workbook, sheetName) {
	return;
}

function refineCancelData(cancelData) {
	let refinedData;
	return refinedData;
}

function createWorkbook() {
	return XlsxPopulate.fromBlankAsync();
}
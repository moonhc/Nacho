function createExcel(cancelData) {
	return createWorkbook().then(
		function (workbook) {
			createCancelSheets(workbook, cancelData);
		});
}

function createCancelSheets(workbook, cancelData) {
	let refinedData = refineCancelData(cancelData);
	writeToSheet(workbook, )
	return workbook;
}

function writeToSheet(workbook, data) {
	return workbook;
}

function downloadExcel(workbook) {
	workbook.outputAsync().then(
		function (blob) {
			let url = window.URL.createObjectURL(blob);
			let a = document.createElement('a');
			document.body.appendChild(a);
			a.href = url;
			a.download = 'out.xlsx';
			a.click();
			window.URL.revokeObjecURL(url);
			document.body.removeChild(a);
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
function createCancelExcel(cancelData) {
	return createWorkbook().then(
		function (workbook) {
			createCancelSheets(workbook, cancelData);
			return workbook;
		});
}

function createCancelSheets(workbook, cancelData) {
	let refinedData = refineCancelData(cancelData);

	for (let type in refinedData) {
		let sheetName = `${type}(취소)`;
		const sheet = workbook.addSheet(sheetName);
		writeToCancelSheet(sheet, refinedData[type]);
	}
}

function writeToCancelSheet(sheet, data) {
	// Write header first
	sheet.range(1, 1, 1, data.range.maxCol)
		.value([data.header]);

	// Write data
	sheet.range(2, 1, data.range.maxRow, data.range.maxCol)
		.value(data.data);
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
			window.URL.revokeObjectURL(url);
			document.body.removeChild(a);
		}).catch(
		function (err) {
			alert(err.message || err);
		});
}

function findSheetByType(workbook, PGType) {
	return;
}

function refineCancelData(cancelData) {
	let refinedData = {};

	for (let id in cancelData) {
		const cancelArray = cancelData[id];
		const type = cancelArray.PGType;

		if (!(type in refinedData)) {
			let headers = Object.keys(cancelArray[0].row);

			refinedData[type] = {
				range: {maxRow: 1, maxCol: headers.length},
				header: headers,
				data: []
			};
		}

		for (let data of cancelArray) {
			refinedData[type].data.push(Object.values(data.row));
			refinedData[type].range.maxRow++;
		}
	}
	return refinedData;
}

function createWorkbook() {
	return XlsxPopulate.fromBlankAsync();
}
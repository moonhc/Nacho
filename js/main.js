const STATE = ['WAIT', 'PROCESSING', 'DONE'];

const rABS = typeof FileReader !== "undefined" && typeof FileReader.prototype !== "undefined" && typeof FileReader.prototype.readAsBinaryString !== "undefined";

let inputFile = false;
let outputFile;

let wb;

let statusIdx = 0;
console.log('[STATE]', STATE[statusIdx]);

const statusSpan = document.querySelector('span.status');
const dotSpan = document.querySelector('span.dot');

const processBtn = document.querySelector('input.process');
const downloadBtn = document.querySelector('input.download');

const fileInput = document.querySelector('input.fileUpload');

processBtn.addEventListener('click', processBtnHandler);
downloadBtn.addEventListener('click', downloadBtnHandler);

fileInput.addEventListener('change', fileInputHandler);

function processBtnHandler() {
	// Check input uploaded
	if (inputFile) {
		// Get the excel file and start to process it
		outputFile = false;
		downloadBtn.disabled = true;
		statusIdx = 1;
		console.log('[STATE]', STATE[statusIdx]);
		statusSpan.innerText = STATE[statusIdx];
		let intervalID =setInterval(dotProcessing, 100);
		startProcessing(intervalID);
	}
	else {
		alert("Upload an excel file first.");
	}
}

function downloadBtnHandler() {
	// Check output made
	if (outputFile) {
		// Return ouput
		console.log("Download complete.");
	}
}

function fileInputHandler(e) {
	let file = fileInput.files[0];

	if (fileValidation(file)) {
		let reader = new FileReader();
		inputFile = true;

		reader.onload = function(e) {
			let data = rABS ? e.target.result : btoa( fixdata(e.target.result) );
			wb = XLSX.read(data, {type: rABS ? 'binary' : 'base64'});
		}

		if(rABS) reader.readAsBinaryString(file);
		else reader.readAsArrayBuffer(file);
	}
	else {
		alert("The file is not a supported format.");
	}

}

function fileValidation(file) {
	let valSet = ['.xlsx', '.xls'];

	for (let val of valSet) {
		if (file.name.match(val))
			return true;
	}

	return false;
}

function startProcessing(intervalID) {
	init();
	// Parse
	for (let payType of PAYTYPE) {
		if (payType == 'ERROR') continue;

		let parserFunc = parserDeterminant(payType);
		parserFunc(wb);
	}

    let rawWB = createRawDataExcel(rawData);
    rawWB.then(function (workbook) {downloadExcel(workbook, 'data.xlsx')});

	// Make a cancel sheet
	let = promiseWB = createCancelExcel(cancelData);
	promiseWB.then(function (workbook) {downloadExcel(workbook, '취소내역.xlsx')});	

	setTimeout(
		function () { 
			statusIdx = 2;
			console.log('[STATE]', STATE[statusIdx]);
			statusSpan.innerText = STATE[statusIdx];

			outputFile = true;
			downloadBtn.disabled = false;
			clearInterval(intervalID);
			dotSpan.innerText = '';
		}, 1000);
}

function dotProcessing() {
	if (dotSpan.innerText.length < 5) {
		dotSpan.innerText += '.';
	}
	else {
		dotSpan.innerText = '.';
	}
}

function fixdata(data) {
	var o = "", l = 0, w = 10240;
	for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
	o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
	return o;
}

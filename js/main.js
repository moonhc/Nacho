const STATE = ['WAIT', 'PROCESSING', 'DONE'];

let inputFile = true;
let outputFile;

let statusIdx = 0;
console.log('[STATE]', STATE[statusIdx]);

const statusSpan = document.querySelector('span.status');
const dotSpan = document.querySelector('span.dot');

const processBtn = document.querySelector('input.process');
const downloadBtn = document.querySelector('input.download');

processBtn.addEventListener('click', processBtnHandler);
downloadBtn.addEventListener('click', downloadBtnHandler);

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
		console.log("Download complete");
	}
}

function startProcessing(intervalID) {
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
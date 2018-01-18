// 결제 종류
const PAYTYPE = [
	// Not handled data
	'ERROR',

	// 국내 온라인 결제
	'INIPAY',
	'ALLAT',
	'DOUZONE',

	// 국외 온라인 결제
	'EXIMBAY',

	// 계좌이체
	'TRANSFER',

	// 현장등록
	'ONSITE'
	];

// Parser functions
const PARSER = [];

// Parsed data from several excel files.
// {등록번호 : {총입금액, 총 수수료 (수수료+부가세), 실입금액 (총입금액-총 수수료), PG사, 차수, 환율} }
let rawData = {};

let errLog = [];

// Distributor for parser
function parserDeterminant(payType) {
	if (payType in PARSER)
		return PARSER[payType];
	else
		return PARSER['ERROR'];

}

// Parser functions
function parserError(sheetName, cellPos, errorType) {
	switch(errorType) {
		case 'dupNo':
			errLog.push(
				`[${errorType}] 시트: ${sheetName} 위치: ${cellPos.r}, ${cellPos.c}
								중복된 등록번호가 있습니다.`);
			break;
		case 'emptyCell':
			errLog.push(
				`[${errorType}] 시트: ${sheetName} 위치: ${cellPos.r}, ${cellPos.c}
								필요한 정보가 없습니다.`);
			break;
		default:
			errLog.push(
				`[${errorType}] 시트: ${sheetName} 위치: ${cellPos.r}, ${cellPos.c}
								인식되지 않은 오류입니다.`);
			break;
	}
}

function parserInipay(wb) {
	let ws = wb.Sheets['이니시스'];
	let rowArr = XLSX.utils.sheet_to_json(ws);
	let rowNum = 1;
	for (let row of rowArr) {
		rowNum++;
		let id = row['주문번호'];

		if (id in rawData) { 
			parserError('이니시스', {r: rowNum, c: '주문번호'}, 'dupNo');
			continue;
		}

		let tmp = {};
		
		let totalFee = parseFloat( row['거래금액'].replace(/[,.]/g, '') );
		if (!totalFee) {
			parserError('이니시스', {r: rowNum, c: '거래금액'}, 'emptyCell');
			continue;
		}
		tmp['총입금액'] = totalFee;

		let tax1 = parseFloat( row['수수료'].replace(/[,.]/g, '') );
		if (!tax1) {
			parserError('이니시스', {r: rowNum, c: '수수료'}, 'emptyCell');
			continue;
		}
		let tax2 = parseFloat( row['부가세'].replace(/[,.]/g, '') );
		if (!tax2) {
			parserError('이니시스', {r: rowNum, c: '부가세'}, 'emptyCell');
			continue;
		}
		tmp['총수수료'] = tax1 + tax2;

		let realFee = parseFloat( row['지급액'].replace(/[,.]/g, '') );
		if (!realFee) {
			parserError('이니시스', {r: rowNum, c: '지급액'}, 'emptyCell');
			continue;
		}
		tmp['실입금액'] = realFee;

		let PGType = '이니시스';
		tmp['PG사'] = PGType;

		tmp['치수'] = '';

		tmp['환율'] = '';

		rawData[id] = tmp;
	}
}

function parserAllat(wb) {
	let ws = wb.Sheets['올앳샘플'];
	let rowArr = XLSX.utils.sheet_to_json(ws);
	let rowNum = 1;
	for (let row of rowArr) {
		rowNum++;
		let id = row['주문번호'];

		if (id in rawData) { 
			parserError('올앳샘플', {r: rowNum, c: '주문번호'}, 'dupNo');
			console.log(id);
			continue;
		}

		let tmp = {};
		
		let totalFee = parseFloat( row['정산금액'].replace(/[,.]/g, '') );
		if (!totalFee) {
			parserError('올앳샘플', {r: rowNum, c: '정산금액'}, 'emptyCell');
			continue;
		}
		tmp['총입금액'] = totalFee;

		let tax1 = parseFloat( row['수수료'].replace(/[,.]/g, '') );
		if (!tax1) {
			parserError('올앳샘플', {r: rowNum, c: '수수료'}, 'emptyCell');
			continue;
		}
		let tax2 = parseFloat( row['수수료부가세'].replace(/[,.]/g, '') );
		if (!tax2) {
			parserError('올앳샘플', {r: rowNum, c: '수수료부가세'}, 'emptyCell');
			continue;
		}
		tmp['총수수료'] = tax1 + tax2;

		let realFee = parseFloat( row['실입금액'].replace(/[,.]/g, '') );
		if (!realFee) {
			parserError('올앳샘플', {r: rowNum, c: '실입금액'}, 'emptyCell');
			continue;
		}
		tmp['실입금액'] = realFee;

		let PGType = '올앳샘플';
		tmp['PG사'] = PGType;

		tmp['치수'] = '';

		tmp['환율'] = '';

		rawData[id] = tmp;
	}
}

function parserDouzone(wb) {
	let ws = wb.Sheets['더존샘플'];
	let rowArr = XLSX.utils.sheet_to_json(ws);

}

function parserEximbay(wb) {
	let ws = wb.Sheets["이니시스"];
	let rowArr = XLSX.utils.sheet_to_json(ws);
	for (let row of rowArr) {
		
	}
}

function parserTransfer(wb) {

}

function parserOnsite(wb) {

}

function init() {
	for (let type of PAYTYPE) {
		funcName = 'parser' + type[0] + type.slice(1).toLowerCase();
		PARSER[type] = window[funcName];
	}
}

init();

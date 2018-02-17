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
let PARSER = {};

// Parsed data from several excel files.
// {
//	등록번호 : 
//		{총입금액USD/KRW, 총수수료(수수료+부가세)USD/KRW, 실입금액(총입금액-총 수수료)KRW, 
//			PG사, 차수, 환율} 
// }
let rawData, cancelData, errLog, mergedData, bankData;

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
		case 'notNumber':
			errLog.push(
				`[${errorType}] 시트: ${sheetName} 위치: ${cellPos.r}, ${cellPos.c}
								숫자로 변환될 수 없습니다.`);
			break;
		case 'noPayment':
			errLog.push(
				`[${errorType}] 시트: ${sheetName} 위치: ${cellPos.r}, ${cellPos.c}
								취소 내역에 대한 결제 내역을 찾을 수 없습니다.`);
			break;
		case 'cancelError':
			errLog.push(
				`[${errorType}] 시트: ${sheetName} 위치: ${cellPos.r}, ${cellPos.c}
								취소 내역이 다른 결제사의 결제 내역의 등록번호와 일치합니다.`);
			break;
		default:
			errLog.push(
				`[${errorType}] 시트: ${sheetName} 위치: ${cellPos.r}, ${cellPos.c}
								인식되지 않은 오류입니다.`);
			break;
	}
}

function parserInipay(wb) {
	let sheetName = '이니시스'
	if (!(sheetName in wb.Sheets)) return;

	let ws = wb.Sheets[sheetName];
	let rowArr = XLSX.utils.sheet_to_json(ws);
	let rowNum = 1;

	for (let row of rowArr) {
		rowNum++;

		let id = row['주문번호'];
		if (!id) {
			parserError(sheetName, {r: rowNum, c: '주문번호'}, 'emptyCell');
			continue;
		}

		let totalFee = row['거래금액'];
		if (!totalFee) {
			parserError(sheetName, {r: rowNum, c: '거래금액'}, 'emptyCell');
			continue;
		}
		totalFee = totalFee.replace(/[,]/g, '');
		if (isNaN(totalFee)) {
			parserError(sheetName, {r: rowNum, c: '거래금액'}, 'notNumber');
			continue;
		}
		totalFee = parseFloat(totalFee);

		let tax1 = row['수수료'];
		if (!tax1) {
			parserError(sheetName, {r: rowNum, c: '수수료'}, 'emptyCell');
			continue;
		}
		tax1 = tax1.replace(/[,]/g, '');
		if (isNaN(tax1)) {
			parserError(sheetName, {r: rowNum, c: '수수료'}, 'notNumber');
			continue;
		}
		tax1 = parseFloat(tax1);

		let tax2 = row['부가세'];
		if (!tax2) {
			parserError(sheetName, {r: rowNum, c: '부가세'}, 'emptyCell');
			continue;
		}
		tax2 = tax2.replace(/[,]/g, '');
		if (isNaN(tax2)) {
			parserError(sheetName, {r: rowNum, c: '부가세'}, 'notNumber');
			continue;
		}
		tax2 = parseFloat(tax2);

		let realFee = totalFee - tax1 - tax2;
		let PGType = '이니시스';

		addData(
			sheetName,
			{r:rowNum, c:'주문번호'},
			{id:id, totalFee:totalFee, tax:tax1+tax2, 
				realFee:realFee, PGType:PGType, currency:undefined},
			row);
	}
}

function parserAllat(wb) {
	let sheetName = '올앳샘플'
	if (!(sheetName in wb.Sheets)) return;
	
	let ws = wb.Sheets[sheetName];
	let rowArr = XLSX.utils.sheet_to_json(ws);
	let rowNum = 1;

	for (let row of rowArr) {
		rowNum++;
		let id = row['주문번호'];
		if (!id) {
			parserError(sheetName, {r: rowNum, c: '주문번호'}, 'emptyCell');
			continue;
		}

		let totalFee = row['정산금액'];
		if (!totalFee) {
			parserError(sheetName, {r: rowNum, c: '정산금액'}, 'emptyCell');
			continue;
		}
		totalFee = totalFee.replace(/[,]/g, '');
		if (isNaN(totalFee)) {
			parserError(sheetName, {r: rowNum, c: '정산금액'}, 'notNumber');
			continue;
		}
		totalFee = parseFloat(totalFee);

		let tax1 = row['수수료'];
		if (!tax1) {
			parserError(sheetName, {r: rowNum, c: '수수료'}, 'emptyCell');
			continue;
		}
		tax1 = tax1.replace(/[,]/g, '');
		if (isNaN(tax1)) {
			parserError(sheetName, {r: rowNum, c: '수수료'}, 'notNumber');
			continue;
		}
		tax1 = parseFloat(tax1);

		let tax2 = row['수수료부가세'];
		if (!tax2) {
			parserError(sheetName, {r: rowNum, c: '수수료부가세'}, 'emptyCell');
			continue;
		}
		tax2 = tax2.replace(/[,]/g, '');
		if (isNaN(tax2)) {
			parserError(sheetName, {r: rowNum, c: '수수료부가세'}, 'notNumber');
			continue;
		}
		tax2 = parseFloat(tax2);

		let realFee = totalFee - tax1 - tax2;
		let PGType = '올앳';
		
		addData(
			sheetName,
			{r:rowNum, c:'주문번호'},
			{id:id, totalFee:totalFee, tax:tax1+tax2, 
				realFee:realFee, PGType:PGType, currency:undefined},
			row);
	}
}

function parserDouzone(wb) {
	let sheetName = '더존샘플'
	if (!(sheetName in wb.Sheets)) return;
	
	let ws = wb.Sheets[sheetName];
	let rowArr = XLSX.utils.sheet_to_json(ws);
	let rowNum = 1;

	for (let row of rowArr) {
		rowNum++;

		if (rowNum % 2 == 1) {
			let id = row['상품명'].replace(/.*\[(\w*)\]/g,'$1');
			if (!id) {
				parserError(sheetName, {r: rowNum, c: '주문번호'}, 'emptyCell');
				continue;
			}
			id = id.replace(/.*\[(\w*)\]/g,'$1');
			
			let totalFee = row['매입금액'];
			if (!totalFee) {
				parserError(sheetName, {r: rowNum, c: '매입금액'}, 'emptyCell');
				continue;
			}
			totalFee = totalFee.replace(/[,]/g, '');
			if (isNaN(totalFee)) {
				parserError(sheetName, {r: rowNum, c: '매입금액'}, 'notNumber');
				continue;
			}
			totalFee = parseFloat(totalFee);

			let tax1 = row['수수료'];
			if (!tax1) {
				parserError(sheetName, {r: rowNum, c: '수수료'}, 'emptyCell');
				continue;
			}
			tax1 = tax1.replace(/[,]/g, '');
			if (isNaN(tax1)) {
				parserError(sheetName, {r: rowNum, c: '수수료'}, 'notNumber');
				continue;
			}
			tax1 = parseFloat(tax1);

			let tax2 = row['부가세'];
			if (!tax2) {
				parserError(sheetName, {r: rowNum, c: '부가세'}, 'emptyCell');
				continue;
			}
			tax2 = tax2.replace(/[,]/g, '');
			if (isNaN(tax2)) {
				parserError(sheetName, {r: rowNum, c: '부가세'}, 'notNumber');
				continue;
			}
			tax2 = parseFloat(tax2);

			let realFee = totalFee - tax1 - tax2;
			let PGType = '더존';

			addData(
				sheetName, 
				{r:rowNum, c:'상품명'},
				{id:id, totalFee:totalFee, tax:tax1+tax2, 
					realFee:realFee, PGType:PGType, currency:undefined},
				row);
		}
	}
}

function parserEximbay(wb) {
	let sheetName = '엑심베이'
	if (!(sheetName in wb.Sheets)) return;
	
	let ws = wb.Sheets[sheetName];
	let rowArr = XLSX.utils.sheet_to_json(ws);
	let rowNum = 1;

	for (let row of rowArr) {
		rowNum++;

		let id = row['등록번호'];
		if (!id) {
			parserError(sheetName, {r: rowNum, c: '등록번호'}, 'emptyCell');
			continue;
		}

		let totalFee = row['금액'];
		if (!totalFee) {
			parserError(sheetName, {r: rowNum, c: '금액'}, 'emptyCell');
			continue;
		}
		totalFee = totalFee.replace(/[,]/g, '');
		if (isNaN(totalFee)) {
			parserError(sheetName, {r: rowNum, c: '금액'}, 'notNumber');
			continue;
		}
		totalFee = parseFloat(totalFee);

		let tax1 = row['수수료'];
		if (!tax1) {
			parserError(sheetName, {r: rowNum, c: '수수료'}, 'emptyCell');
			continue;
		}
		tax1 = tax1.replace(/[,]/g, '');
		if (isNaN(tax1)) {
			parserError(sheetName, {r: rowNum, c: '수수료'}, 'notNumber');
			continue;
		}
		tax1 = parseFloat(tax1);

		let tax2 = row['결제수수료'];
		if (!tax2) {
			parserError(sheetName, {r: rowNum, c: '결제수수료'}, 'emptyCell');
			continue;
		}
		tax2 = tax2.replace(/[,]/g, '');
		if (isNaN(tax2)) {
			parserError(sheetName, {r: rowNum, c: '결제수수료'}, 'notNumber');
			continue;
		}
		tax2 = parseFloat(tax2);

		let realFee = totalFee + tax1 + tax2;
		let PGType = '엑심베이';

		let currency = row['환율']
		if (!currency) {
			parserError(sheetName, {r: rowNum, c: '환율'}, 'emptyCell');
			continue;
		}
		currency = currency.replace(/[,]/g, '');
		if (isNaN(currency)) {
			parserError(sheetName, {r: rowNum, c: '환율'}, 'notNumber');
			continue;
		}
		currency = parseFloat(currency);

		addData(
			sheetName, 
			{r:rowNum, c:'등록번호'},
			{id:id, totalFee:totalFee, tax:(tax1+tax2)*-1, 
				realFee:realFee*currency, PGType:PGType, currency:currency},
			row);
	}
}

function parserTransfer(wb) {
	let sheetName = '계좌이체'
	if (!(sheetName in wb.Sheets)) return;
	
    let ws = wb.Sheets[sheetName]
    let rowArr = XLSX.utils.sheet_to_json(ws);
    let rowNum = 1;

    for (let row of rowArr) {
        rowNum++;

        let id = row['등록번호'];
        let origin = row['출처'];
        if (!id && !origin) {
            continue;
        } else if (id in rawData) {
            parserError(sheetName, {r: rowNum, c: '등록번호'}, 'dupNo');
            continue;
        }

        let totalFee = row['맡기신금액'];
        if (!totalFee) {
            parserError(sheetName, {r: rowNum, c: '맡기신금액'}, 'emptyCell');
            continue;
        }
        totalFee = totalFee.replace(/[,]/g, '');
        if (isNaN(totalFee)) {
        	parserError(sheetName, {r: rowNum, c: '맡기신금액'}, 'notNumber');
        	continue;
        }
        totalFee = parseFloat(totalFee);

        if (id) {
		    let tmp = {};

		    tmp['총입금액'] = totalFee;

		    tmp['총수수료'] = 0;

		    tmp['실입금액'] = totalFee;

		    let PGType = '계좌이체';
		    tmp['PG사'] = PGType;

		    rawData[id] = tmp;
		}
		else if (origin) {
	        if (row['출처'] in bankData) {
	        	bankData[row['출처']] += totalFee;
	        } else if (!row['출처']) {
	        	bankData[row['출처']] = totalFee;
	        }
	    }
    }
}

function parserOnsite(wb) {
	let sheetName = '현장카드'
	if (!(sheetName in wb.Sheets)) return;
	
    let ws = wb.Sheets[sheetName]
    let rowArr = XLSX.utils.sheet_to_json(ws);
    let rowNum = 1;
    let prevId = 0;

    for (let row of rowArr) {
        rowNum++;

        let totalFee = row['카드금액'];
        if (!totalFee) {
        	parserError(sheetName, {r: rowNum, c: '카드금액'}, 'emptyCell');
        	continue;
        }
        totalFee = totalFee.replace(/[,]/g, '');
        if (isNaN(totalFee)) {
        	parserError(sheetName, {r: rowNum, c: '카드금액'}, 'notNumber');
        	continue;
        }
        totalFee = parseFloat(totalFee);

        let tax = row['차감수수료'];
        if (!tax) {
        	parserError(sheetName, {r: rowNum, c: '차감수수료'}, 'emptyCell');
        	continue;
        }
        tax = tax.replace(/[,]/g, '');
        if (isNaN(tax)) {
        	parserError(sheetName, {r: rowNum, c: '차감수수료'}, 'notNumber');
        	continue;
        }
        tax = parseFloat(tax);

		let realFee = row['입금액'];
        if (!realFee) {
            parserError(sheetName, {r: rowNum, c: '입금액'}, 'emptyCell');
            continue;
        }
        realFee = realFee.replace(/[,]/g, '');
        if (isNaN(realFee)) {
        	parserError(sheetName, {r: rowNum, c: '입금액'}, 'notNumber');
        	continue;
        }
        realFee = parseFloat(realFee);

        let id = row['등록번호'];
        if (id in rawData) {
            parserError(sheetName, {r: rowNum, c: '등록번호'}, 'dupNo');
            continue;
        } else if (!id) {
            if (!row['카드금액']) {
                parserError(sheetName, {r: rowNum, c: '등록번호'}, 'emptyCell');
                continue;
            } else {
                rawData[prevId]['총입금액'] += totalFee;
                rawData[prevId]['총수수료'] += tax;
                rawData[prevId]['실입금액'] += realFee;
                continue;
            }
        }

        let tmp = {};
     
        tmp['총입금액'] = totalFee;
        tmp['총수수료'] = tax;
        tmp['실입금액'] = realFee;

        let PGType = '현장카드';
        tmp['PG사'] = PGType;

        rawData[id] = tmp;
        prevId = id;
    }
}

function addData(sheetName, cellPos, data, row) {
	let id = data.id;
	let totalFee = data.totalFee;
	let tax = data.tax;
	let realFee = data.realFee;
	let PGType = data.PGType
	let currency = data.currency;
	if (totalFee >= 0) {
		// 결제 내역
		if (id in rawData) { 
			parserError(sheetName, cellPos, 'dupNo');
			return;
		}

		let tmp = {};
		tmp['총입금액'] = totalFee;
		tmp['총수수료'] = tax;
		tmp['실입금액'] = realFee;
		tmp['PG사'] = PGType;
		tmp['환율'] = currency;
		tmp['row'] = row;

		rawData[id] = tmp;
	} else if (totalFee < 0) {
		// 취소 내역
		if (!(id in rawData)) { 
			parserError(sheetName, cellPos, 'noPayment');
			return;
		}

		if (id in cancelData) {
			if (cancelData[id].PGType == PGType) {
				let tmp = {};
				tmp['총입금액'] = totalFee;
				tmp['총수수료'] = tax;
				tmp['실입금액'] = realFee;
				tmp['PG사'] = PGType;
				tmp['환율'] = currency;
				tmp['row'] = row;
				cancelData[id].push(tmp);
			} else {
				parserError(sheetName, cellPos, 'cancelError');
				return;
			}
		} else {
			let originData = JSON.parse(JSON.stringify(rawData[id]));

			if (originData['PG사'] == PGType) {
				cancelData[id] = [];
				
				cancelData[id]['PGType'] = PGType;

				// Original Data
				cancelData[id].push(originData);

				// Cancel Data
				let tmp = {};
				tmp['총입금액'] = totalFee;
				tmp['총수수료'] = tax;
				tmp['실입금액'] = realFee;
				tmp['PG사'] = PGType;
				tmp['환율'] = currency;
				tmp['row'] = row;
				cancelData[id].push(tmp); 
			} else {
				parserError(sheetName, cellPos, 'cancelError');
				return;
			}
		}

		rawData[id]['총입금액'] += totalFee;
		rawData[id]['총수수료'] += tax;
		rawData[id]['실입금액'] += realFee;

		if (rawData[id]['총입금액'] == 0) {
			delete rawData[id];
		}
	} else {
		// Error
		parserError(sheetName, cellPos, 'Undefined');
		return;
	}
}

function createRawDataExcel(data) {
    return createWorkbook().then(
        function (workbook) {
        	createRawDataSheet(workbook, data);
            return workbook;
        });
}

function createRawDataSheet(workbook, data) {
	let mergedData = mergeToRawData(data);
	let headers = Object.keys(mergedData[0]);
	let outputData = {
	    range: {maxRow: 1, maxCol: headers.length},
	    header: headers,
	    data: []
	};

	for (let data in mergedData) {
	    outputData.data.push(Object.values(mergedData[data]));
	    outputData.range.maxRow++;
	}
	//console.log(outputData);

	const sheet = workbook.addSheet('RawData');
	//sheet.range(1, 1, 1, outputData.range.maxCol).value([outputData.header]);
	sheet.range(1, 1, outputData.range.maxRow, outputData.range.maxCol).value(outputData.data);
}

function mergeToRawData(data) {
    let sheetName = 'RawData'
    let ws = wb.Sheets[sheetName];
    let rowArr = XLSX.utils.sheet_to_json(ws, {header: 1, raw: true, defval: ''});
    let rowNum = -1;
    let output = {};

    rowArr[0].push('총입금액', '총수수료', '실입금액', 'PG사', '환율');
    for (let row of rowArr) {
        rowNum++;

        let id = row[7];
		if (!id) {
            output[rowNum] = row;
			continue;
		}

        let tmp = row;
        for (let key in data[id]) {
            if (key != 'row') {
                tmp.push(data[id][key]);
            }
        }

        output[rowNum] = tmp;
    }

    mergedData = {}
    for (let i=1;i<Object.keys(output).length;i++) {
    	mergedData[i-1] = {}
    	for (let j=0;j<output[0].length;j++) {
    		mergedData[i-1][output[0][j]] = output[i][j]
		}
	}

    return output;
}
            
function init() {
	for (let type of PAYTYPE) {
		funcName = 'parser' + type[0] + type.slice(1).toLowerCase();
		PARSER[type] = window[funcName];
	}

	rawData = {};
	cancelData = {};
	errLog = [];
	bankData = {};
}

init();

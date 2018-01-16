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
let rawData = [];

// Distributor for parser
function parserDeterminant(payType) {
	if (payType in PARSER)
		return PARSER[payType];
	else
		return PARSER['ERROR'];

}

// Parser functions
function parserError() {

}

function parserInipay() {

}

function parserAllat() {

}

function parserDouzone() {

}

function parserEximbay() {

}

function parserTransfer() {

}

function parserOnsite() {

}

function init() {
	for (let type of PAYTYPE) {
		funcName = 'parser' + type[0] + type.slice(1).toLowerCase();
		PARSER[type] = window[funcName];
	}
}

init();

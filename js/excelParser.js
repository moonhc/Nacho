// Not handled data
const ERROR		= 0;
// 국내 온라인 결제
const INIPAY	= 1;
const ALLAT		= 2;
const DOUZONE	= 3;
// 국외 온라인 결제
const EXIMBAY	= 4;
// 계좌이체
const TRANSFER	= 5;
// 현장등록
const ONSITE	= 6;

// 결제 종류
const PAYMENT_TYPE = [ERROR, INIPAY, ALLAT, DOUZONE, EXIMBAY, TRANSFER, ONSITE];

// Parsed data from several excel files.
// {총입금액, 총 수수료 (수수료+부가세), 실입금액 (총입금액-총 수수료), PG사, 차수, 환율}
let rawData = [];

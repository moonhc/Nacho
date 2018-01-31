// let XLSX = require('xlsx')
// let XlsxPopulate = require('xlsx-populate')
// let errLog = []

let output = null
let errRow = {}

// Parser functions
function parserErrorForProcessing(sheetName, regNum, errorType) {
    switch(errorType) {
        case 'depositError':
            errLog.push(
            `[${errorType}] 시트: ${sheetName} 등록번호: ${regNum}
								실입금액을 확인할 필요가 있습니다.`);
            break;
        case 'dupNum':
            errLog.push(
                `[${errorType}] 시트: ${sheetName} 등록번호: ${regNum}
								중복된 등록번호가 있습니다.`);
            break;
        case 'emptyCell':
            errLog.push(
                `[${errorType}] 시트: ${sheetName} 등록번호: ${regNum}
								필요한 정보가 없습니다.`);
            break;
        case 'currencyInconsistency':
            errLog.push(
                `[${errorType}] 시트: ${sheetName} 등록번호: ${regNum}
								통화가 실입금액과 일치하지 않습니다.`);
            break;
        default:
            errLog.push(
                `[${errorType}] 시트: ${sheetName} 등록번호: ${regNum}
								인식되지 않은 오류입니다.`);
            break;
    }
}

// let workbook = XLSX.readFile("data (4).xls");
// let rawDataSheet = workbook.Sheets["RawData"];

function getInitStat() {
    return {"KRW":{"Card":{"count":0, "KRW":0, "실입금액":0}, "Transfer":{"count":0, "KRW":0}, "Cash":{"count":0, "KRW":0}},
        "USD":{"Card":{"count":0, "USD":0, "실입금액":0}, "Transfer":{"count":0, "USD":0, "실입금액":0}, "Cash":{"count":0, "USD":0, "실입금액":0}}};
}

function analyze(input, output) {

    function abs(x) {
        return x > 0 ? x : -x
    }

    function searchHeadersStat(headers) {
        let c = filters.length+1;
        while(outputS.cell(1, c).value()) {
            let found = true;
            for(let r=0;r<headers.length;r++) {
                if(outputS.cell(r+1,c).value().indexOf(headers[r]) === -1) {
                    c += 1;
                    found = false;
                    break;
                }
            }
            if(found)
                return c;
        }
        return -1;
    }

    function searchHeadersRaw(header) {
        let c = 1
        while(outputS.cell(1, c).value()) {
            if(outputS.cell(1, c).value() === header) {
                return c
            }
            c += 1
        }
        return -1;
    }

    function getFilters() {
        let c = 1;
        let filters = []
        while (outputS.cell(1, c).value() === undefined) {
            filters.push(outputS.cell(2,c).value());
            c += 1;
        }
        return filters;
    }

    function getTotalFilters() {
        let r = [];
        let c = filters.length + 1;
        while(true) {
            cell = outputS.cell(1, c);
            if(cell.value() === "총합계") {
                r.push(outputS.cell(2, c).value());
            } else if(cell.value() === undefined) {
                break;
            }
            c += 1;
        }
        return r;
    }

    function getFees() {
        let fees = {}
        let r = 3
        while(outputS.cell(r, 1).value() !== "Total") {
            if(outputS.cell(r, 1).value() === outputS.cell(r, 2).value()
            || outputS.cell(r, 2).value() === undefined) {
                fees[[outputS.cell(r, 1).value()+" Count"]] = parseInt(outputS.cell(r, filters.length + 1).value())
            } else {
                fees[Array.apply(null, {length: filters.length}).map(Number.call, Number)
                    .map(function(x){return x+1})
                    .map(function(x){return outputS.cell(r, x).value()})]
                    = parseInt(outputS.cell(r, filters.length + 1).value())

            }
            r += 1
        }
        return fees
    }

    let outputS = output.sheet(0);
    let filters = getFilters();
    let totalFilters = getTotalFilters();
    let outputR = 3;
    let data = input;
    // let data = XLSX.utils.sheet_to_json(input)
    let fees = getFees()

    while(outputS.cell(outputR, 1).value() !== "Total") {
        let stat = getInitStat();

        let curRegFee = parseInt(outputS.cell(outputR, filters.length+1).value())

        let additional = false;
        let tmp = {};
        for(let i=1;i<filters.length+1;i++) {
            let v = outputS.cell(outputR, i).value();
            if(v) {
                tmp[v] = 0;
            } else {
                additional = true;
                tmp["Count"] = 0;
                break;
            }
        }

        let conditions = Object.keys(tmp);
        let inputR = 1;
        let curRow = data[parseInt(inputR)];

        while(curRow["Num"]) {
            if(outputR === 3) {
                if (errRow[curRow["Num"]] !== undefined && errRow[curRow["Num"]].indexOf("dupNum")) {
                    parserErrorForProcessing("RawData", curRow["Num"], "dupNum")
                    errRow[curRow["Num"]].push("dupNum")
                } else if (errRow[curRow["Num"]] === undefined) {
                    errRow[curRow["Num"]] = []
                }
            }
            if(parseFloat(curRow["실입금액"]) && parseFloat(curRow["실입금액"]) > 0) {
                let sumFee = 0
                let feeKeys = Object.keys(fees)
                let picked = []
                // console.log(curRow["Num"], curRow["실입금액"])
                for(let i=0;i<feeKeys.length;i++) {
                    let curRowKeys = Object.keys(curRow)
                    let tmpKeys = feeKeys[i].split(",")
                    if(tmpKeys.length > 1) {
                        let find = Array.apply(null, {length: tmpKeys.length}).map(function (x) {
                            return false
                        })
                        for (let k = 0; k < tmpKeys.length; k++) {
                            for (let j = 0; j < curRowKeys.length; j++) {
                                if (curRow[curRowKeys[j]] === tmpKeys[k]) {
                                    find[k] = true
                                    break
                                }
                            }
                        }
                        if(find.reduce((prev, curr)=> prev && curr)) {
                            picked.push(feeKeys[i])
                            sumFee += fees[feeKeys[i]]
                        }
                    } else {
                        tmpKeys = feeKeys[i].split(" ")
                        for(let j=0;j<curRowKeys.length;j++) {
                            let find = true
                            for(let k=0;k<tmpKeys.length;k++) {
                               if(curRowKeys[j].indexOf(tmpKeys[k]) === -1) {
                                   find = false
                                   break
                               }
                           }
                           if(find) {
                               let count = parseInt(curRow[curRowKeys[j]]) ? parseInt(curRow[curRowKeys[j]]) : 0
                               sumFee += count*fees[feeKeys[i]]
                               picked.push(feeKeys[i])
                               break
                           }
                        }
                    }
                }

                if((curRow["Currency"] === "USD" && curRow["환율"] === undefined && curRow["Pay Method"] !== "Transfer")
                || curRow["Currency"] === "KRW" && curRow["환율"] !== undefined) {
                    parserErrorForProcessing("RawData", curRow["Num"], "currencyInconsistency")
                }

                if(additional) {
                    let t_key = Object.keys(curRow)
                        .filter(function(x) {
                            return conditions.map(function(c) {return x.indexOf(c) !== -1})
                                        .reduce((prev, curr) => prev && curr)})[0]
                    let count = parseInt(curRow[t_key]) ? parseInt(curRow[t_key]) : 0

                    let tFee = count * curRegFee

                    stat[curRow["Currency"]][curRow["Pay Method"]]["count"] += count;
                    if (curRow["Pay Method"] === "Card" || curRow["Currency"] === "USD") {
                        // To do : 실입금액 and 입금액(USD)
                        stat[curRow["Currency"]][curRow["Pay Method"]]["실입금액"] += parseFloat(curRow["실입금액"]) * (tFee/sumFee)
                        if(outputR === 3) {
                            let rate = parseFloat(curRow["환율"]) ? parseFloat(curRow["환율"]) : 1100
                            let shouldGet = sumFee * rate
                            if (abs(shouldGet - parseInt(curRow["실입금액"])) > shouldGet * 0.1 && errRow[curRow["Num"]].indexOf("depositError") === -1) {
                                parserErrorForProcessing("RawData", curRow["Num"], "depositError")
                                errRow[curRow["Num"]].push("depositError")
                            }
                        }
                    }
                } else {
                    tmp = {};
                    for (let i = 0; i < filters.length; i++) {
                        tmp[curRow[filters[i]]] = 0;
                    }
                    let conditions_ = Object.keys(tmp).join();
                    if (conditions_ === conditions.join()) {
                        stat[curRow["Currency"]][curRow["Pay Method"]]["count"] += 1;

                        if (curRow["Pay Method"] === "Card" || curRow["Currency"] === "USD") {
                            // To do : 실입금액 and 입금액(USD)
                            stat[curRow["Currency"]][curRow["Pay Method"]]["실입금액"] += parseFloat(curRow["실입금액"]) * (curRegFee/sumFee)
                            if(outputR === 3) {
                                let rate = parseFloat(curRow["환율"]) ? parseFloat(curRow["환율"]) : 1100
                                let shouldGet = sumFee * rate
                                if (abs(shouldGet - parseInt(curRow["실입금액"])) > shouldGet * 0.1 && errRow[curRow["Num"]].indexOf("depositError") === -1) {
                                    parserErrorForProcessing("RawData", curRow["Num"], "depositError")
                                    errRow[curRow["Num"]].push("depositError")
                                }
                            }
                        }
                    }
                }
            }
            inputR += 1;
            curRow = data[inputR];
        }

        for(let currency in stat) {
            let regFee = outputS.cell(outputR, searchHeadersStat([currency])).value();
            for(let pm in stat[currency]) {
                let t_idx = searchHeadersStat([pm, currency]);
                outputS.cell(outputR, t_idx - 1).value(stat[currency][pm]["count"]);
                outputS.cell(outputR, t_idx).value(stat[currency][pm]["count"]*regFee);
                if(pm === "Card" || currency === "USD") {
                    outputS.cell(outputR, t_idx+1).value(stat[currency][pm]["실입금액"]);
                }
            }
        }
        for(let item in totalFilters) {
            let t_idx = searchHeadersStat(["총합계", totalFilters[item]])
            let c = filters.length + 3;
            let s = 0;
            while(true) {
                if(outputS.cell(2, c).value() === undefined
                    || outputS.cell(1, c).value() === "총합계") {
                    break;
                } else if(outputS.cell(2, c).value() === totalFilters[item]) {
                    let tmp = parseInt(outputS.cell(outputR, c).value());
                    if(tmp)
                        s += tmp
                }
                c += 1;
            }
            outputS.cell(outputR, t_idx).value(s);
        }
        outputR += 1;
    }

    let c = filters.length + 3;
    while(true) {
        if(outputS.cell(2, c).value() === undefined)
            break;
        outputS.cell(outputR, c).formula("=SUM("+outputS.cell(3,c).address()+":"+outputS.cell(outputR-1,c).address()+")")
        c += 1;
    }
    return output
}
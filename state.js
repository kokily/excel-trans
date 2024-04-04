"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var xlsx = require("xlsx");
var iconv_1 = require("iconv");
var fs = require("fs");
var jschartdet = require("jschardet");
var excelDir = './data';
function csvToJson(target) {
    var unEncoding = fs.readFileSync(target);
    var encoding = jschartdet.detect(unEncoding).encoding;
    var iconv = new iconv_1.Iconv(encoding, 'utf-8');
    var result = iconv.convert(unEncoding).toString('utf-8');
    var source = xlsx.read(result, { type: 'binary' });
    var sheet = source.Sheets[source.SheetNames[0]];
    var jsonData = xlsx.utils.sheet_to_json(sheet, {
        raw: true,
        header: [
            '순번',
            '입고일자',
            '품번',
            '품명',
            '규격',
            '수량',
            '단위',
            '단가',
            '금액',
            '부가세',
            '합계금액',
            '거래처명',
            '적요',
            '특이사항',
            '현장명',
            'PJT코드',
            'PJT명',
            '입고창고',
        ],
    });
    jsonData.shift();
    jsonData.pop();
    return jsonData;
}
function menufactureJson(target) {
    var sources = target.map(function (item) {
        return {
            구매일자: item.입고일자,
            납품장소: item.현장명,
            품명: item.품명,
            규격: item.규격,
            세: item.부가세 === 0 ? '면' : '과',
            단위: item.단위,
            수량: parseFloat(item.수량),
            단가: parseFloat(item.단가),
            금액: parseFloat(item.합계금액),
            업체명: item.거래처명,
        };
    });
    return sources;
}
function saveToExcel(target, workBook, divide) {
    var sheetData = xlsx.utils.json_to_sheet(target);
    xlsx.utils.book_append_sheet(workBook, sheetData, divide);
}
/**
 * contract: 계약 식자재
 * nonContract: 비계약 식자재
 * employee: 계약 직원식자재
 * nonEmployee: 비계약 직원식자재
 */
function classificationItems(target, workBook) {
    return __awaiter(this, void 0, void 0, function () {
        var prevContract, prevNonContract, contract, nonContract, employee, nonEmployee;
        return __generator(this, function (_a) {
            prevContract = target.filter(function (data) { return data.업체명 === '삼성웰스토리'; });
            prevNonContract = target.filter(function (data) { return data.업체명 === '삼성웰스토리(비)'; });
            contract = prevContract.filter(function (data) { return data.납품장소 !== '직원식당'; });
            nonContract = prevNonContract.filter(function (data) { return data.납품장소 !== '직원식당-비계약'; });
            employee = prevContract.filter(function (data) { return data.납품장소 === '직원식당'; });
            nonEmployee = prevNonContract.filter(function (data) { return data.납품장소 === '직원식당-비계약'; });
            if (contract.length > 0) {
                saveToExcel(contract, workBook, '계약 식자재');
            }
            if (nonContract.length > 0) {
                saveToExcel(nonContract, workBook, '비계약 식자재');
            }
            if (employee.length > 0) {
                saveToExcel(employee, workBook, '계약 직원 식자재');
            }
            if (nonEmployee.length > 0) {
                saveToExcel(nonEmployee, workBook, '비계약 직원 식자재');
            }
            xlsx.writeFile(workBook, "".concat(excelDir, "/\uACB0\uC0B0\uC11C.xlsx"));
            return [2 /*return*/];
        });
    });
}
function bootStrap() {
    return __awaiter(this, void 0, void 0, function () {
        var target, data, workBook, err_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    target = csvToJson("".concat(excelDir, "/state.csv"));
                    data = menufactureJson(target);
                    workBook = xlsx.utils.book_new();
                    return [4 /*yield*/, classificationItems(data, workBook)];
                case 1:
                    _a.sent();
                    return [3 /*break*/, 3];
                case 2:
                    err_1 = _a.sent();
                    console.log(err_1);
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    });
}
bootStrap();

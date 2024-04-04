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
// 사업장 리스트
var MessRoom = [
    '양식뷔페',
    '양식뷔페-비계약',
    '양식소모품',
    '양식소모품-비계약',
    '양정식',
    '양정식-비계약',
    '연회부',
    '연회부-비계약',
    '연회부소모품',
    '연회부소모품-비계약',
    '운영지원부',
    '운영지원부-비계약',
    '중식뷔페',
    '중식뷔페-비계약',
    '중식소모품',
    '중식소모품-비계약',
    '중정식',
    '중정식-비계약',
    '직원식당',
    '직원식당-비계약',
    '한정식',
    '한정식-비계약',
];
var excelDir = './data';
function csvToJson(target) {
    var source = xlsx.readFile(target);
    var sheet = source.Sheets[source.SheetNames[0]];
    var jsonData = xlsx.utils.sheet_to_json(sheet, {
        raw: true,
        header: [
            '날짜',
            '사업장명',
            '코드',
            '단품명',
            '면과세',
            '규격',
            '단위',
            '원산지',
            '검수수량',
            '단가',
            '금액',
            '행사구분',
            '거래명세서 분리구분',
        ],
    });
    jsonData.shift();
    return jsonData;
}
function workPlace(target) {
    var source = target.split('(')[3].replace(')', '');
    switch (source) {
        case '중식당':
            source = '중정식';
            break;
        case '양식당':
            source = '양정식';
            break;
        case '중식당-비계약':
            source = '중정식-비계약';
            break;
        case '양식당-비계약':
            source = '양정식-비계약';
            break;
        case '연회부-식재료':
            source = '연회부';
            break;
        default:
            break;
    }
    return source;
}
function manufactureJson(target) {
    var sources = target.map(function (item) {
        return {
            사업장명: workPlace(item.사업장명),
            단위: item.단위,
            세: item.면과세 === '면세' ? '면' : '과',
            품목코드: "24".concat(item.코드),
            품목명: item.단품명,
            규격: item.규격,
            바코드: '',
            수량: item.검수수량,
            평균단가: item.단가,
            입고금액: item.금액,
            부가세: item.면과세 === '면세' ? 0 : item.금액 * 0.1,
        };
    });
    return sources;
}
function classificationItems(target, workBook) {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            MessRoom.map(function (mass) {
                var prevData = target.filter(function (data) { return data.사업장명 === mass; });
                if (prevData.length > 0) {
                    var freeTax = prevData.filter(function (data) { return data.세 === '면'; });
                    var taxation = prevData.filter(function (data) { return data.세 === '과'; });
                    if (freeTax.length > 0) {
                        var freeSheet = xlsx.utils.json_to_sheet(freeTax);
                        xlsx.utils.book_append_sheet(workBook, freeSheet, "".concat(mass, "-\uBA74"));
                    }
                    if (taxation.length > 0) {
                        var taxationSheet = xlsx.utils.json_to_sheet(taxation);
                        xlsx.utils.book_append_sheet(workBook, taxationSheet, "".concat(mass, "-\uACFC"));
                    }
                }
            });
            return [2 /*return*/];
        });
    });
}
function bootStrap() {
    return __awaiter(this, void 0, void 0, function () {
        var target, data, workBook, sheetData, err_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    target = csvToJson("".concat(excelDir, "/hyundai.xlsx"));
                    data = manufactureJson(target);
                    workBook = xlsx.utils.book_new();
                    sheetData = xlsx.utils.json_to_sheet(data);
                    xlsx.utils.book_append_sheet(workBook, sheetData, '총괄');
                    return [4 /*yield*/, classificationItems(data, workBook)];
                case 1:
                    _a.sent();
                    xlsx.writeFile(workBook, "".concat(excelDir, "/\uBD84\uC57C\uBCC4.xlsx"));
                    return [3 /*break*/, 3];
                case 2:
                    err_1 = _a.sent();
                    console.error(err_1);
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    });
}
bootStrap();

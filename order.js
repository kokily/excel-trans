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
    '국방컨벤션(양식뷔페-계약)',
    '국방컨벤션(양식뷔페-비계약)',
    '국방컨벤션(양식소모품-계약)',
    '국방컨벤션(양식소모품-비계약)',
    '국방컨벤션(양정식-계약)',
    '국방컨벤션(양정식-비계약)',
    '국방컨벤션(연회부-계약)',
    '국방컨벤션(연회부-비계약)',
    '국방컨벤션(연회부소모품-계약)',
    '국방컨벤션(연회부소모품-비계약)',
    '국방컨벤션(운영지원부-계약)',
    '국방컨벤션(운영지원부-비계약)',
    '국방컨벤션(중식뷔페-계약)',
    '국방컨벤션(중식뷔페-비계약)',
    '국방컨벤션(중식소모품-계약)',
    '국방컨벤션(중식소모품-비계약)',
    '국방컨벤션(중정식-계약)',
    '국방컨벤션(중정식-비계약)',
    '국방컨벤션(직원식당-계약)',
    '국방컨벤션(직원식당-비계약)',
    '국방컨벤션(한정식-계약)',
    '국방컨벤션(한정식-비계약)',
];
var excelDir = './data';
function csvToJson(target) {
    var source = xlsx.readFile(target);
    var sheet = source.Sheets[source.SheetNames[0]];
    var jsonData = xlsx.utils.sheet_to_json(sheet, {
        raw: true,
        header: [
            'MESSROOMNAME',
            'LARGENAME',
            'DAY',
            'MATERIALCODE',
            'GUBUN',
            'SELLERMATERIALCODE',
            'MATERIALNAME',
            'UNITNAME',
            'DIMENSION',
            'STOCKINQTY',
            'PRICE',
            'SELLINGAMT',
            'SELLINGTAX',
            'TOT_AMT',
        ],
    });
    jsonData.shift();
    return jsonData;
}
function manufactureJson(target) {
    var sources = target.map(function (item) {
        return {
            사업장명: item.MESSROOMNAME.replace('/', '-'),
            단위: item.UNITNAME,
            세: item.SELLINGTAX === 0 ? '면' : '과',
            품목코드: "23".concat(item.MATERIALCODE),
            품목명: item.MATERIALNAME,
            규격: item.DIMENSION,
            바코드: '',
            수량: parseFloat(item.STOCKINQTY),
            평균단가: parseFloat(item.PRICE),
            입고금액: parseFloat(item.SELLINGAMT),
            부가세: parseFloat(item.SELLINGTAX),
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
                    target = csvToJson("".concat(excelDir, "/order.csv"));
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
                    console.log(err_1);
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    });
}
bootStrap();

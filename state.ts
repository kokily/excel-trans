import type { WorkBook } from 'xlsx';
import * as xlsx from 'xlsx';
import { Iconv } from 'iconv';
import * as fs from 'fs';
import * as jschartdet from 'jschardet';

// 변환될 양식
type Trans = {
  구매일자: string;
  납품장소: string;
  품명: string;
  규격: string;
  세: string;
  단위: string;
  수량: number;
  단가: number;
  금액: number;
  업체명: string;
};

const excelDir = './data';

function csvToJson(target: string) {
  const unEncoding = fs.readFileSync(target);
  const encoding = jschartdet.detect(unEncoding).encoding;
  const iconv = new Iconv(encoding, 'utf-8');
  const result = iconv.convert(unEncoding).toString('utf-8');

  const source = xlsx.read(result, { type: 'binary' });
  const sheet = source.Sheets[source.SheetNames[0]];

  const jsonData: Array<JSON> = xlsx.utils.sheet_to_json(sheet, {
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

function menufactureJson(target: Array<JSON>) {
  const sources: Array<Trans> = target.map((item: any) => {
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

type Divide =
  | '계약 식자재'
  | '비계약 식자재'
  | '계약 직원 식자재'
  | '비계약 직원 식자재';

function saveToExcel(target: Trans[], workBook: WorkBook, divide: Divide) {
  const sheetData = xlsx.utils.json_to_sheet(target);

  xlsx.utils.book_append_sheet(workBook, sheetData, divide);
}

/**
 * contract: 계약 식자재
 * nonContract: 비계약 식자재
 * employee: 계약 직원식자재
 * nonEmployee: 비계약 직원식자재
 */
async function classificationItems(target: Trans[], workBook: WorkBook) {
  const prevContract = target.filter((data) => data.업체명 === '삼성웰스토리');
  const prevNonContract = target.filter(
    (data) => data.업체명 === '삼성웰스토리(비)'
  );

  const contract = prevContract.filter((data) => data.납품장소 !== '직원식당');
  const nonContract = prevNonContract.filter(
    (data) => data.납품장소 !== '직원식당-비계약'
  );

  const employee = prevContract.filter((data) => data.납품장소 === '직원식당');
  const nonEmployee = prevNonContract.filter(
    (data) => data.납품장소 === '직원식당-비계약'
  );

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

  xlsx.writeFile(workBook, `${excelDir}/결산서.xlsx`);
}

async function bootStrap() {
  try {
    const target = csvToJson(`${excelDir}/state.csv`);
    const data = menufactureJson(target);

    const workBook = xlsx.utils.book_new();

    await classificationItems(data, workBook);

    // console.log('결산서 엑셀파일 생성!');
    // process.exit(1);
  } catch (err: any) {
    console.log(err);
  }
}

bootStrap();

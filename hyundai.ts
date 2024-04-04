import type { WorkBook } from 'xlsx';
import * as xlsx from 'xlsx';

// 변환된 양식
type Trans = {
  사업장명: string;
  단위: string;
  세: string;
  품목코드: string;
  품목명: string;
  규격: string;
  바코드: string;
  수량: number;
  평균단가: number;
  입고금액: number;
  부가세: number;
};

// 사업장 리스트
const MessRoom = [
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

const excelDir = './data';

function csvToJson(target: string) {
  const source = xlsx.readFile(target);
  const sheet = source.Sheets[source.SheetNames[0]];
  const jsonData: Array<JSON> = xlsx.utils.sheet_to_json(sheet, {
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

function workPlace(target: string) {
  let source = target.split('(')[3].replace(')', '');

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

function manufactureJson(target: Array<JSON>) {
  const sources: Array<Trans> = target.map((item: any) => {
    return {
      사업장명: workPlace(item.사업장명),
      단위: item.단위,
      세: item.면과세 === '면세' ? '면' : '과',
      품목코드: `24${item.코드}`,
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

async function classificationItems(target: Trans[], workBook: WorkBook) {
  MessRoom.map((mass) => {
    const prevData = target.filter((data) => data.사업장명 === mass);

    if (prevData.length > 0) {
      const freeTax = prevData.filter((data) => data.세 === '면');
      const taxation = prevData.filter((data) => data.세 === '과');

      if (freeTax.length > 0) {
        const freeSheet = xlsx.utils.json_to_sheet(freeTax);
        xlsx.utils.book_append_sheet(workBook, freeSheet, `${mass}-면`);
      }

      if (taxation.length > 0) {
        const taxationSheet = xlsx.utils.json_to_sheet(taxation);
        xlsx.utils.book_append_sheet(workBook, taxationSheet, `${mass}-과`);
      }
    }
  });
}

async function bootStrap() {
  try {
    const target = csvToJson(`${excelDir}/hyundai.xlsx`);
    const data = manufactureJson(target);

    const workBook = xlsx.utils.book_new();

    const sheetData = xlsx.utils.json_to_sheet(data);

    xlsx.utils.book_append_sheet(workBook, sheetData, '총괄');

    await classificationItems(data, workBook);

    xlsx.writeFile(workBook, `${excelDir}/분야별.xlsx`);
  } catch (err: any) {
    console.error(err);
  }
}

bootStrap();

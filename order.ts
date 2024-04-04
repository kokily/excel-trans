import type { WorkBook } from 'xlsx';
import * as xlsx from 'xlsx';

// 변환될 양식
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

const excelDir = './data';

function csvToJson(target: string) {
  const source = xlsx.readFile(target);
  const sheet = source.Sheets[source.SheetNames[0]];
  const jsonData: Array<JSON> = xlsx.utils.sheet_to_json(sheet, {
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

function manufactureJson(target: Array<JSON>) {
  const sources: Array<Trans> = target.map((item: any) => {
    return {
      사업장명: item.MESSROOMNAME.replace('/', '-'),
      단위: item.UNITNAME,
      세: item.SELLINGTAX === 0 ? '면' : '과',
      품목코드: `23${item.MATERIALCODE}`,
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
    const target = csvToJson(`${excelDir}/order.csv`);
    const data = manufactureJson(target);

    const workBook = xlsx.utils.book_new();

    const sheetData = xlsx.utils.json_to_sheet(data);

    xlsx.utils.book_append_sheet(workBook, sheetData, '총괄');

    await classificationItems(data, workBook);

    xlsx.writeFile(workBook, `${excelDir}/분야별.xlsx`);

    // console.log('발주서 엑셀자료 생성 완료!');
    // process.exit(1);
  } catch (err: any) {
    console.log(err);
  }
}

bootStrap();

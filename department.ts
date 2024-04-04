import type { WorkBook } from 'xlsx';
import * as xlsx from 'xlsx';
import * as jschartdet from 'jschardet';
import { Iconv } from 'iconv';
import fs from 'fs';

// 변환 양식
type Trans = {
  부서명: string;
  성명: string;
};

const excelDir = './data';

function csvToJson(target: string) {
  const unEncoding = fs.readFileSync(target);
  const encoding = jschartdet.detect(unEncoding).encoding;
  const iconv = new Iconv(encoding, 'utf-8');
  const result = iconv.convert(unEncoding).toString('utf-8');

  const source = xlsx.read(result, { type: 'binary' });
  const sheet = source.Sheets[source.SheetNames[0]];

  console.log(sheet);
}

async function bootStrap() {
  try {
    csvToJson(`${excelDir}/department.csv`);
  } catch (err: any) {
    console.log(err);
  }
}

bootStrap();

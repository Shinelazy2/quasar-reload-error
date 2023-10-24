import xlsx from 'xlsx';
import * as Excel from 'exceljs';
import dayjs from 'dayjs';
import { AutoFilter } from 'exceljs';
const sqlite = require('aa-sqlite');
const path = require('path');

const dbPath = path.join(__dirname) + '/../../db.sqlite';

export class commonExcel {
  /**
   * xls, xlsx 파일을 읽고 배열을 보냄
   * @returns [ ... ]
   */
  readExcel = (readExcelPath: string): string[][] => {
    console.time('common-excel - readExcel');
    const workbook = xlsx.readFile(readExcelPath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows: string[][] = xlsx.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: '',
    });
    console.log(
      '🚀 ~ file: common-excel.ts:21 ~ commonExcel ~ rows:',
      rows[1].length
    );
    console.timeEnd('common-excel - readExcel');
    return rows;
  };

  /**
   * Array를 받아 Excel을 DB로 만듬
   * [ ...컬럼명 ]
   * ex) ApiReponse.data[0]
   */
  // createDB = async (tableName: string, columnList: string[]) => {
  //   console.time('common-excel - createDB')
  //   await sqlite.open(dbPath)

  //   const columns = columnList
  //   const createTableQuery = `
  //   CREATE TABLE IF NOT EXISTS ${tableName} (
  //       ${columns.map(column => `"${column}" TEXT`).join(',\n')}
  //     );
  //   `;
  //   await sqlite.run(createTableQuery);
  //   console.timeEnd('common-excel - createDB')
  // }
  createDB = async (tableName: string, columnList: string[]) => {
    console.time('common-excel - createDB');
    await sqlite.open(dbPath);

    const columns = [];
    const columnCounts: { [key: string]: number } = {};

    for (const column of columnList) {
      if (!columnCounts[column]) {
        // 해당 열 이름이 처음 나타난 경우
        columnCounts[column] = 1;
        columns.push(column);
      } else {
        // 이미 나타난 열 이름인 경우
        columnCounts[column]++;
        columns.push(`${column}_${columnCounts[column]}`);
      }
    }

    const createTableQuery = `
      CREATE TABLE IF NOT EXISTS ${tableName} (
          ${columns.map((column) => `"${column}" TEXT`).join(',\n')}
        );
    `;

    await sqlite.run(createTableQuery);
    console.timeEnd('common-excel - createDB');
  };

  insertDB = async (tableName: string, insertData: string[][]) => {
    console.time('common-excel - insertDB');
    // let isSuccess: { message: string }
    try {
      await sqlite.open(dbPath);

      const batchSize = 100;
      const columns = insertData[0];

      // IMPOTANT!!!
      await sqlite.run('BEGIN TRANSACTION');

      for (let i = 1; i < insertData.length; i += batchSize) {
        const batch = insertData.slice(i, i + batchSize);
        const valuesQueries = [];

        for (const row of batch) {
          const rowValues = row.map((value) => `${value}`).join('", "');
          valuesQueries.push(`("${rowValues}")`);
        }

        const insertQuery = `INSERT INTO ${tableName} ("${columns.join(
          '", "'
        )}") VALUES ${valuesQueries.join(', ')}`;
        await sqlite.run(insertQuery);
      }
      await sqlite.run('COMMIT');
    } catch (error) {
      if (error instanceof Error) console.log(`ERROR - ${error.message}`);
    }
    console.timeEnd('common-excel - insertDB');
  };

  objectToExcel = async (query: any) => {
    console.time('common-excel - objectToExcel');
    const workbook = new Excel.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    const sheet2 = workbook.addWorksheet('Sheet2');

    const columns = Object.keys(query[0]);
    sheet.columns = columns.map((col) => ({
      header: col,
      key: col,
      width: 15,
    }));

    sheet2.columns = columns.map((col) => ({
      header: col,
      key: col,
      width: 15,
    }));

    // TODO: Sheet2에서 필터링 할 수 있게 만들어야함.
    query.forEach((row: any) => {
      sheet2.addRow(row);
    });

    // Sheet2에 테이블 생성
    // const sheet2Table = sheet2.addTable({
    //   name: 'MyTable',
    //   ref: `A1:${String.fromCharCode(65 + columns.length - 1)}${query.length + 1}`;,
    //   headerRow: true,
    // });

    for (let i = 0; i < query.length; i++) {
      const columns = query[i];
      const row = sheet.getRow(i + 2);
      for (let j = 0; j < Object.keys(columns).length; j++) {
        row.getCell(j + 1).value = Object.values(columns)[j] as string;
      }
    }

    const autoFilterRange = `A1:${String.fromCharCode(
      65 + columns.length - 1
    )}${query.length + 1}`;
    sheet.autoFilter = autoFilterRange;
    sheet.addTable;

    await workbook.xlsx.writeFile(
      `C:\\Users\\shine\\OneDrive\\문서\\울산_Test\\${dayjs().format(
        'YYYYMMDDHHmmss'
      )}objectToExcel.xlsx`
    );
    console.log('object to Excel Complete');
    console.timeEnd('common-excel - objectToExcel');
  };

  dropTable = async (tableName: string) => {
    console.time('common-excel - dropTable');
    await sqlite.open(dbPath);

    const dropQuery = `
    DROP TABLE "${tableName}"
    `;
    console.log('🚀 ~ file: ipcDB.ts:718 ~ im.handle ~ dropQuery:', dropQuery);
    const data = await sqlite.get_all(dropQuery, []);
    // return data.message === 'success'
    console.timeEnd('common-excel - dropTable');
    return data.messages === 'success';
  };
}

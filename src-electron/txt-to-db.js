const fs = require('fs');
const sqlite = require('aa-sqlite');
const path = require('path');
const iconv = require('iconv-lite');

async function txtToDb(txtPath) {
  console.time('txtToDB');

  const dbPath = path.join(__dirname, '../../db.sqlite');
  await sqlite.open(dbPath);

  convertANSItoUTF8(txtPath);

  const dropQuery = `
    DROP TABLE IF EXISTS "tmpTXT"
  `;
  console.log('txtToDb - INFO : tmpTXT table drop');
  await sqlite.get(dropQuery, []);

  const data = await fs.promises.readFile(txtPath, 'utf-8');

  const columnCounts = new Set();
  data.split('\n').forEach((line) => {
    const columns = line.split('\t');
    columnCounts.add(columns.length);
  });

  const columnNames = Array.from(
    { length: columnCounts.values().next().value },
    (_, i) => `column${i + 1}`
  );
  const createTableQuery = `
    CREATE TABLE IF NOT EXISTS "tmpTXT" (${columnNames
      .map((name) => `${name} TEXT`)
      .join(', ')})
  `;
  await sqlite.get(createTableQuery);

  for (const line of data.split('\n')) {
    const columns = line.split('\t').map((column) => column.trim());
    const values = columns.map((column) => `"${column}"`).join(', ');
    const insertQuery = `
      INSERT INTO "tmpTXT" (${columnNames.join(', ')}) VALUES (${values})
    `;
    await sqlite.get(insertQuery);
  }
  console.timeEnd('txtToDB');
  return true;
}

function convertANSItoUTF8(fileName) {
  const content = fs.readFileSync(fileName);

  const bom = content.slice(0, 3);
  let encoding;
  if (bom.equals(Buffer.from([0xef, 0xbb, 0xbf]))) {
    encoding = 'utf-8';
  } else if (bom.equals(Buffer.from([0xfe, 0xff]))) {
    encoding = 'utf-16be';
  } else if (bom.equals(Buffer.from([0xff, 0xfe]))) {
    encoding = 'utf-16le';
  } else {
    encoding =
      iconv.decode(content, 'ISO-8859-1') ===
      iconv.decode(content, 'Windows-1252')
        ? 'ISO-8859-1'
        : 'EUC-KR';
  }

  if (encoding !== 'EUC-KR') {
    const utf8Content = iconv.decode(content, 'EUC-KR');
    fs.writeFileSync(fileName, utf8Content);
  }
}

module.exports = txtToDb;

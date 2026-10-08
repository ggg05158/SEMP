import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const input = new URL('../.qa/equipment-list-verify.xlsx', import.meta.url);
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(input.pathname.replace(/^\/(\w:)/, '$1')));
console.log((await workbook.inspect({kind:'table', range:'장비목록!A1:K4', tableMaxRows:4, tableMaxCols:11, maxChars:3500})).ndjson);
console.log((await workbook.inspect({kind:'match', searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!', options:{useRegex:true,maxResults:20},maxChars:1000})).ndjson);
for (const [sheetName, range, filename] of [
    ['장비목록','A1:I4','equipment-list-preview.png'],
    ['사용 안내','A1:A13','equipment-list-guide.png']]) {
  const image = await workbook.render({sheetName, range, scale:1.5});
  await fs.writeFile(new URL('../.qa/'+filename,import.meta.url),new Uint8Array(await image.arrayBuffer()));
}

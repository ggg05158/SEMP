import fs from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import {Workbook, SpreadsheetFile} from '@oai/artifact-tool';

const out = new URL('../outputs/semp/', import.meta.url);
const wb = Workbook.create();
const sheet = wb.worksheets.add('장비등록');
sheet.showGridLines = false;
sheet.getRange('A1:D1').values = [['장치이름', '일련번호 처음', '일련번호 끝', '용도']];
sheet.getRange('A1:C101').format.font = {name: 'Malgun Gothic', size: 11, color: '#26354A'};
sheet.getRange('A1:C101').format.rowHeight = 25;
sheet.getRange('A1:C1').format = {fill: '#16344A', font: {name: 'Malgun Gothic', size: 11, bold: true, color: '#FFFFFF'}, rowHeight: 32};
sheet.getRange('A1:C1').format.horizontalAlignment = 'center';
sheet.getRange('A2:C101').format.fill = '#F2F8F6';
sheet.getRange('A1:A101').format.columnWidth = 32;
sheet.getRange('B1:C101').format.columnWidth = 25;
sheet.getRange('A2:C101').setNumberFormat('@');
sheet.getRange('D1:D101').copyFrom(sheet.getRange('C1:C101'), 'all');
sheet.getRange('D1').values = [['용도']];
sheet.getRange('D1:D101').format.columnWidth = 20;
sheet.getRange('D1:D101').format.font = {name:'Malgun Gothic', size:11, color:'#26354A'};
sheet.getRange('D1').format = {fill:'#16344A', font:{name:'Malgun Gothic',size:11,bold:true,color:'#FFFFFF'}, horizontalAlignment:'center'};
sheet.getRange('D2:D101').format.fill = '#F2F8F6';
sheet.getRange('D2:D101').setNumberFormat('@');
sheet.freezePanes.freezeRows(1);
sheet.getRange('F2').values = [['SEMP 장비 등록']];
sheet.getRange('F2').format.font = {name: 'Malgun Gothic', size: 16, bold: true, color: '#16344A'};
sheet.getRange('F4:F11').values = [
 ['왼쪽 네 열에 2행부터 입력하세요.'],
 ['처음과 끝: 양 끝을 포함해 일련번호 생성'],
 ['처음만: 개별 장비 1개 생성'],
 ['둘 다 비움: 장치만 등록'],
 ['접두어 범위 예: PSU-001 ~ PSU-010'],
 ['끝만 입력하거나 역순인 범위는 사용할 수 없습니다.'],
 ['일련번호는 텍스트 형식을 유지하세요. 앞자리 0이 보존됩니다.'],
 ['한 범위 최대 10,000개, 한 파일 최대 20,000개 일련번호']
];
sheet.getRange('F4:F24').format.font = {name:'Malgun Gothic', size:11, color:'#5A6C80'};
sheet.getRange('F1:F24').format.columnWidth = 78;
sheet.getRange('F4:F24').format.rowHeight = 25;
sheet.getRange('F13:F17').values = [
 ['용도 안내'], ['기본 용도: 비행용, 예비용, 환경용, 점검용'],
 ['다른 용도를 직접 입력해도 됩니다. 등록하면 용도 목록에 추가됩니다.'],
 ['한 행의 용도는 해당 행에서 생성되는 모든 일련번호에 적용됩니다.'],
 ['장치만 등록할 때는 용도를 비워두세요.']
];
sheet.getRange('F13').format.font.bold = true;
sheet.getRange('F19').values = [['입력 예시 (왼쪽에 필요한 값만 작성)']];
sheet.getRange('F19').format.font = {name:'Malgun Gothic',size:11,bold:true,color:'#16344A'};
sheet.getRange('F20:F23').values = [
 ['통신장치 A / 001 / 020 / 비행용'], ['전원장치 B / PSU-001 / PSU-010 / 예비용'],
 ['측정장치 C / SN-105 / (비움) / 점검용'], ['케이블 D / (비움) / (비움) / (비움)']
];
wb.recalculate();
console.log((await wb.inspect({kind:'table',range:'장비등록!A1:D5',include:'values',tableMaxRows:5,tableMaxCols:4})).ndjson);
console.log((await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!',options:{useRegex:true,maxResults:10}})).ndjson);
await fs.mkdir(out,{recursive:true});
const preview=await wb.render({sheetName:'장비등록',range:'A1:F24',scale:1.1,format:'png'});
await fs.writeFile(new URL('../../.qa/template.png',out),new Uint8Array(await preview.arrayBuffer()));
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(fileURLToPath(new URL('SEMP_장비등록_템플릿.xlsx',out)));

import {test} from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {parseCSV,mergeCSV,encodeCSV} from './csv.js';
test('four attached files merge into 4000 rows and round-trip without data loss',()=>{
  const files=[1,2,3,4].map(n=>({name:`${n}.csv`,rows:parseCSV(readFileSync(new URL(`./samples/이름_점수_${n}.csv`,import.meta.url),'utf8'))}));
  const rows=mergeCSV(files); assert.equal(rows.length,4001); assert.deepEqual(rows[0],['이름','점수']);
  assert.deepEqual(parseCSV(encodeCSV(rows)),rows);
  for(let i=0;i<4;i++) assert.deepEqual(rows.slice(1+i*1000,1+(i+1)*1000),files[i].rows.slice(1));
});
test('quoted commas, newlines, quotes, BOM and CRLF',()=>{
  const rows=[['이름','메모'],['김,하나','줄1\n줄2'],['이"름','']];
  assert.deepEqual(parseCSV(encodeCSV(rows)),rows);
});
test('rejects malformed files and mismatched headers',()=>{
  for(const s of ['', '이름,점수\n홍길동', 'a\n"unclosed', 'a\n"closed"bad']) assert.throws(()=>parseCSV(s));
  assert.throws(()=>mergeCSV([{name:'a',rows:[['이름','점수']]},{name:'b',rows:[['점수','이름']]}]));
});

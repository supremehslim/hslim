export function parseCSV(text) {
  text = text.replace(/^\uFEFF/, '');
  const rows = []; let row = [], cell = '', quoted = false, closed = false;
  const push = () => { row.push(cell); if (row.some(v => v !== '')) rows.push(row); row = []; cell = ''; closed = false; };
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') { cell += '"'; i++; }
      else if (c === '"') { quoted = false; closed = true; }
      else cell += c;
    } else if (c === ',' ) { row.push(cell); cell = ''; closed = false; }
    else if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; push(); }
    else if (c === '"' && cell === '' && !closed) quoted = true;
    else { if (closed || c === '"') throw new Error('CSV 따옴표 형식을 확인해 주세요.'); cell += c; }
  }
  if (quoted) throw new Error('닫히지 않은 따옴표가 있어요.');
  if (cell !== '' || row.length || closed) push();
  if (!rows.length) throw new Error('비어 있는 CSV 파일이에요.');
  if (rows.some(r => r.length !== rows[0].length)) throw new Error('행마다 열 개수가 달라요.');
  return rows;
}
export function mergeCSV(files) {
  if (!files.length) throw new Error('파일을 먼저 선택해 주세요.');
  const header = files[0].rows[0];
  for (const f of files) if (JSON.stringify(f.rows[0]) !== JSON.stringify(header)) throw new Error(`${f.name}: 열 이름이나 순서가 달라요.`);
  return [header, ...files.flatMap(f => f.rows.slice(1))];
}
export function encodeCSV(rows) {
  return '\uFEFF' + rows.map(r => r.map(c => /[",\r\n]/.test(c) ? '"' + c.replaceAll('"', '""') + '"' : c).join(',')).join('\r\n') + '\r\n';
}

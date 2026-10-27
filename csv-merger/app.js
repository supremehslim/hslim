import { parseCSV, mergeCSV, encodeCSV } from './csv.js';
const $ = id => document.getElementById(id);
let files = [], merged = null, busy = false;
const format = n => n.toLocaleString('ko-KR');
function invalidate() { merged = null; $('result').hidden = true; $('result-empty').hidden = false; $('download').disabled = true; $('result-badge').textContent = '준비 중'; }
function render() {
  $('files').replaceChildren();
  files.forEach((f, i) => {
    const li = document.createElement('li');
    const icon = document.createElement('span'); icon.className = 'file-icon'; icon.textContent = 'CSV';
    const info = document.createElement('div'); info.className = 'file-info';
    const name = document.createElement('strong'); name.textContent = f.name;
    const size = document.createElement('small'); size.textContent = `${format(f.rows.length - 1)}행 · ${(f.size / 1024).toFixed(1)} KB`;
    info.append(name, size);
    const remove = document.createElement('button'); remove.className = 'remove'; remove.textContent = '×'; remove.setAttribute('aria-label', `${f.name} 삭제`); remove.disabled = busy;
    remove.onclick = () => { files.splice(i, 1); invalidate(); render(); $('status').textContent = ''; };
    li.append(icon, info, remove); $('files').append(li);
  });
  $('empty-files').hidden = !!files.length;
  $('file-count').textContent = `${files.length}개 선택`;
  $('merge').disabled = busy || !files.length; $('clear').disabled = busy || !files.length; $('samples').disabled = busy; $('file-input').disabled = busy;
}
async function addFiles(incoming) {
  if (busy) return;
  busy = true; render(); $('status').textContent = '파일을 읽고 있어요…';
  const errors = [];
  for (const f of incoming) {
    try {
      if (!f.name.toLowerCase().endsWith('.csv')) throw new Error('CSV 파일만 선택해 주세요.');
      const bytes = await f.arrayBuffer(); let text;
      try { text = new TextDecoder('utf-8', {fatal: true}).decode(bytes); } catch { text = new TextDecoder('euc-kr', {fatal: true}).decode(bytes); }
      files.push({name:f.name, size:f.size, rows:parseCSV(text)});
    } catch (e) { errors.push(`${f.name}: ${e.message}`); }
  }
  files.sort((a,b) => a.name.localeCompare(b.name,'ko',{numeric:true}));
  invalidate(); busy = false; render(); $('status').textContent = errors.join(' / ');
}
$('file-input').onchange = async e => { await addFiles([...e.target.files]); e.target.value = ''; };
for (const event of ['dragover', 'dragleave', 'drop']) $('dropzone').addEventListener(event, e => {
  e.preventDefault(); $('dropzone').classList.toggle('drag', event === 'dragover');
  if (event === 'drop') addFiles([...e.dataTransfer.files]);
});
$('samples').onclick = async () => {
  if (busy) return;
  busy = true; render(); $('status').textContent = '첨부 파일을 불러오는 중이에요…';
  try {
    const incoming = await Promise.all([1,2,3,4].map(async n => { const name = `이름_점수_${n}.csv`; const response = await fetch(`samples/${encodeURIComponent(name)}`); if (!response.ok) throw new Error('첨부 파일을 불러오지 못했어요.'); return new File([await response.blob()], name); }));
    files = []; busy = false; await addFiles(incoming);
  } catch(e) { $('status').textContent = e.message; busy = false; render(); }
};
$('clear').onclick = () => { files = []; invalidate(); render(); $('status').textContent = ''; };
$('merge').onclick = () => {
  try {
    merged = mergeCSV(files);
    $('total-files').textContent = `${files.length}개`; $('total-rows').textContent = `${format(merged.length - 1)}행`; $('total-cols').textContent = `${merged[0].length}개`;
    const table = $('preview'); table.replaceChildren();
    merged.slice(0,11).forEach((row,i) => { const tr = document.createElement('tr'); [i === 0 ? 'No.' : String(i), ...row].forEach(value => { const cell = document.createElement(i === 0 ? 'th' : 'td'); cell.textContent = value; tr.append(cell); }); table.append(tr); });
    $('preview-note').textContent = `전체 ${format(merged.length - 1)}행 중 처음 ${Math.min(10,merged.length - 1)}행을 표시합니다.`;
    $('result').hidden = false; $('result-empty').hidden = true; $('download').disabled = false; $('result-badge').textContent = '합치기 완료'; $('status').textContent = '';
  } catch(e) { invalidate(); $('status').textContent = e.message; }
};
$('download').onclick = () => {
  if (!merged) return;
  const url = URL.createObjectURL(new Blob([encodeCSV(merged)], {type:'text/csv;charset=utf-8'}));
  const a = document.createElement('a'); a.href = url; a.download = '이름_점수_통합.csv'; a.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
};
render();

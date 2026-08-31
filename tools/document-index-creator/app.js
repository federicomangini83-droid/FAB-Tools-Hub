let currentArtifact = null;
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
const input = document.getElementById('files');
bindDrop('files', 'drop');
input.onchange = () => {
  document.getElementById('list').textContent = [...input.files].map(f => f.name).join(' · ');
};
async function parsePdf(f) {
  const pdf = await pdfjsLib.getDocument({ data: await f.arrayBuffer() }).promise;
  const out = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const text = content.items.map(x => x.str).join(' ').replace(/\s+/g, ' ').trim();
    if (text) out.push({ source: f.name, type: 'pdf', section: `Page ${p}`, content: text });
  }
  return out;
}
async function parseDocx(f) {
  const r = await mammoth.extractRawText({ arrayBuffer: await f.arrayBuffer() });
  const blocks = r.value.split(/\n\s*\n/).map(x => x.replace(/\s+/g, ' ').trim()).filter(Boolean);
  return blocks.map((content, i) => ({ source: f.name, type: 'docx', section: `Block ${i + 1}`, content }));
}
function outputs(rows, name) {
  const json = JSON.stringify(rows, null, 2);
  const csv = Papa.unparse(rows);
  const txt = rows.map(r => `[${r.source} | ${r.section}]\n${r.content}`).join('\n\n');
  return { json, csv, txt, name };
}
document.getElementById('process').onclick = async () => {
  try {
    const fs = [...input.files];
    if (!fs.length) throw new Error('Select at least one DOCX or PDF file.');
    setStatus(`Processing ${fs.length} file(s)...`);
    let rows = [];
    for (const f of fs) {
      if (f.name.toLowerCase().endsWith('.pdf')) rows.push(...await parsePdf(f));
      else if (f.name.toLowerCase().endsWith('.docx')) rows.push(...await parseDocx(f));
      else throw new Error(`Unsupported format: ${f.name}`);
    }
    if (!rows.length) throw new Error('No extractable text found. Scanned PDFs may require OCR.');
    const name = safeName(document.getElementById('name').value);
    const o = outputs(rows, name);
    const format = document.getElementById('format').value;
    if (format === 'zip') {
      const z = new JSZip();
      z.file(name + '.json', o.json); z.file(name + '.csv', o.csv); z.file(name + '.txt', o.txt);
      const blob=await z.generateAsync({type:'blob'});currentArtifact={name:name+'.zip',blob};saveAs(blob,name+'.zip');
    } else {
      const mime = format === 'json' ? 'application/json' : format === 'csv' ? 'text/csv' : 'text/plain';
      const blob=new Blob([o[format]],{type:mime+';charset=utf-8'});currentArtifact={name:name+'.'+format,blob};saveAs(blob,currentArtifact.name);
    }
    storage.artifactReady();setStatus(`Completed: ${rows.length} index records created.`, 'ok');
  } catch (e) { setStatus(e.message, 'err'); }
};
document.getElementById('reset').onclick = () => location.reload();

const storage=initGitHubStorage({tool:'document-index-creator',folder:'storage/document-index',getArtifact:()=>currentArtifact});

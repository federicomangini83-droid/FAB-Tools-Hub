pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';

let currentArtifact = null;

const input = document.getElementById('files');
const formatSelect = document.getElementById('format');

bindDrop('files', 'drop');

input.addEventListener('change', () => {
  document.getElementById('list').textContent = [...input.files]
    .map(file => file.name)
    .join(' · ');
});

async function parsePdf(file) {
  const pdf = await pdfjsLib.getDocument({
    data: await file.arrayBuffer()
  }).promise;

  const rows = [];

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    const text = content.items
      .map(item => item.str)
      .join(' ')
      .replace(/\s+/g, ' ')
      .trim();

    if (text) {
      rows.push({
        source: file.name,
        type: 'pdf',
        section: `Page ${pageNumber}`,
        document_content: text
      });
    }
  }

  return rows;
}

async function parseDocx(file) {
  const result = await mammoth.extractRawText({
    arrayBuffer: await file.arrayBuffer()
  });

  const blocks = result.value
    .split(/\n\s*\n/)
    .map(text => text.replace(/\s+/g, ' ').trim())
    .filter(Boolean);

  return blocks.map((text, index) => ({
    source: file.name,
    type: 'docx',
    section: `Block ${index + 1}`,
    document_content: text
  }));
}

function createChunk(row, index) {
  return {
    id: String(index + 1),
    source: row.source,
    type: row.type,
    section: row.section,
    document_content: row.document_content
  };
}

function paddedId(index, total) {
  const width = Math.max(4, String(total).length);
  return String(index + 1).padStart(width, '0');
}

async function createJsonChunksZip(chunks, outputName, includeConsolidated) {
  const zip = new JSZip();
  const folder = zip.folder('json');

  chunks.forEach((chunk, index) => {
    const id = paddedId(index, chunks.length);
    folder.file(
      `${outputName}_${id}.json`,
      JSON.stringify(chunk, null, 2)
    );
  });

  if (includeConsolidated) {
    zip.file(`${outputName}.json`, JSON.stringify(chunks, null, 2));
    zip.file(`${outputName}.csv`, Papa.unparse(chunks));
    zip.file(
      `${outputName}.txt`,
      chunks
        .map(chunk =>
          `[${chunk.id} | ${chunk.source} | ${chunk.section}]\n${chunk.document_content}`
        )
        .join('\n\n')
    );
  }

  return zip.generateAsync({ type: 'blob' });
}

function createSingleOutput(chunks, format) {
  if (format === 'json') {
    return {
      extension: 'json',
      mime: 'application/json',
      content: JSON.stringify(chunks, null, 2)
    };
  }

  if (format === 'csv') {
    return {
      extension: 'csv',
      mime: 'text/csv',
      content: Papa.unparse(chunks)
    };
  }

  return {
    extension: 'txt',
    mime: 'text/plain',
    content: chunks
      .map(chunk =>
        `[${chunk.id} | ${chunk.source} | ${chunk.section}]\n${chunk.document_content}`
      )
      .join('\n\n')
  };
}

document.getElementById('process').addEventListener('click', async () => {
  try {
    const files = [...input.files];

    if (!files.length) {
      throw new Error('Select at least one DOCX or PDF file.');
    }

    setStatus(`Processing ${files.length} file(s)...`);

    let extractedRows = [];

    for (const file of files) {
      const lowerName = file.name.toLowerCase();

      if (lowerName.endsWith('.pdf')) {
        extractedRows.push(...await parsePdf(file));
      } else if (lowerName.endsWith('.docx')) {
        extractedRows.push(...await parseDocx(file));
      } else {
        throw new Error(`Unsupported format: ${file.name}`);
      }
    }

    if (!extractedRows.length) {
      throw new Error(
        'No extractable text found. Scanned PDFs may require the OCR backend.'
      );
    }

    const chunks = extractedRows.map(createChunk);
    const outputName = safeName(document.getElementById('name').value);
    const format = formatSelect.value;

    if (format === 'json-zip' || format === 'all') {
      setStatus(`Creating ZIP with ${chunks.length} JSON file(s)...`);

      const blob = await createJsonChunksZip(
        chunks,
        outputName,
        format === 'all'
      );

      currentArtifact = {
        name: `${outputName}.zip`,
        blob
      };
    } else {
      const output = createSingleOutput(chunks, format);
      currentArtifact = {
        name: `${outputName}.${output.extension}`,
        blob: new Blob([output.content], {
          type: `${output.mime};charset=utf-8`
        })
      };
    }

    saveAs(currentArtifact.blob, currentArtifact.name);
    storage.artifactReady();

    setStatus(
      `Completed: ${chunks.length} JSON chunk(s) extracted. Output: ${currentArtifact.name}`,
      'ok'
    );
  } catch (error) {
    currentArtifact = null;
    storage.artifactReady();
    setStatus(error.message, 'err');
  }
});

document.getElementById('reset').addEventListener('click', () => {
  location.reload();
});

const storage = initGitHubStorage({
  tool: 'document-index-creator',
  folder: 'storage/document-index',
  getArtifact: () => currentArtifact
});

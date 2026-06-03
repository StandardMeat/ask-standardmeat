// claude-proxy/index.js — Ask Standard Meat retrieval proxy
// Phase 1 (metadata-only index): model-driven tool-use loop.
// Tools: search_files (metadata only) and read_file (one file's text on demand).
//
// IMPORTANT: heavy libraries (mammoth, xlsx, pdf-parse, @azure/storage-blob) are
// required INSIDE the functions that use them, never at module top level. Per the
// build-index startup lesson (AZFD0005 / "node exited with code 1" / gRPC 14
// UNAVAILABLE), heavy top-level requires run in the worker init window and can crash
// the host<->worker handshake on cold start.

const { getGraphToken, getSiteId } = require('../shared/graph');

const MAX_TOOL_ITERATIONS = 5;

// Folders whose contents must never be read or surfaced (sensitive PII).
// Stopgap guardrail; the durable fix is removing/locking these files in SharePoint.
const RESTRICTED_PATHS = [
    'shabaka documents/admin/dload/'
];

const STOP_WORDS = new Set(['that','this','with','from','find','show','what','have','will','where','when','which','about','your','they','them','there','their','would','could','should','please','refer','look','tell','give','make','need','want','help','using','sends','pulling','scripts','script','file','files','process','actually','looking','supposed','then','another','how','many','are','in','on','at','of','an','the','do','does','did','was','were','been','being','can','also','any','all','some','our']);

const EXPANSIONS = {
    'accounts payable': ['ap'], 'accounts receivable': ['ar'], 'general ledger': ['gl'],
    'purchase order': ['po'], 'sales order': ['so'], 'inventory': ['inv'],
    'payable': ['ap'], 'receivable': ['ar'], 'vendor': ['vend'],
    'ap': ['payable','accounts'], 'ar': ['receivable','accounts'], 'gl': ['ledger'],
    'po': ['purchase'], 'edi': ['edi']
};

const TOOLS = [
    {
        name: 'search_files',
        description: 'Search the Standard Meat SharePoint document library by keyword. Matches against file names, paths, and folder names. Returns up to 15 matching files as metadata only: name, path, folder, file extension, and an openable SharePoint link. To read the actual contents of a file, call read_file with the file path (the "path" field from these results).',
        input_schema: {
            type: 'object',
            properties: {
                query: {
                    type: 'string',
                    description: 'Keywords to search for. Examples: "accounts payable vendor export", "OnBase purchase order merge", or a specific script name.'
                }
            },
            required: ['query']
        }
    },
    {
        name: 'read_file',
        description: 'Read the text contents of a single file from the SharePoint library, given its exact path as returned by search_files. Supports Word (.docx), Excel (.xlsx/.xls), PDF, and plain-text files. Returns up to ~30,000 characters of extracted text. Call this after search_files, before answering questions about what a file actually contains.',
        input_schema: {
            type: 'object',
            properties: {
                path: {
                    type: 'string',
                    description: 'The exact file path to read, copied from the "path" field of a search_files result.'
                }
            },
            required: ['path']
        }
    }
];

let cachedIndex = null;

async function loadIndex(context) {
    if (cachedIndex) return cachedIndex;
    try {
        const { BlobServiceClient } = require('@azure/storage-blob');
        const conn = process.env.AzureWebJobsStorage;
        const service = BlobServiceClient.fromConnectionString(conn);
        const blobClient = service.getContainerClient('index').getBlobClient('file-index.json');
        const buffer = await blobClient.downloadToBuffer();
        const parsed = JSON.parse(buffer.toString('utf8'));
        cachedIndex = Array.isArray(parsed.files) ? parsed.files : [];
        context.log('Index loaded:', cachedIndex.length, 'files');
        return cachedIndex;
    } catch (e) {
        context.log('Index load failed:', e.message);
        return null;
    }
}

function isRestricted(path) {
    const p = String(path || '').toLowerCase();
    return RESTRICTED_PATHS.some(function (r) { return p.includes(r); });
}

function searchFiles(query, files) {
    const text = String(query || '').toLowerCase();
    let keywords = (text.match(/[a-z_]{2,}/g) || []).filter(function (w) { return !STOP_WORDS.has(w) && w.length >= 2; });

    const expanded = new Set(keywords);
    for (const kw of keywords) {
        if (EXPANSIONS[kw]) EXPANSIONS[kw].forEach(function (e) { expanded.add(e); });
    }
    for (const phrase in EXPANSIONS) {
        if (phrase.includes(' ') && text.includes(phrase)) EXPANSIONS[phrase].forEach(function (e) { expanded.add(e); });
    }
    keywords = Array.from(expanded).filter(function (w) { return w.length >= 2; });
    if (keywords.length === 0) return [];

    const visible = files.filter(function (f) { return !isRestricted(f.path); });
    const scored = visible
        .map(function (f) {
            const name = (f.name || '').toLowerCase();
            const path = (f.path || '').toLowerCase();
            const folder = (f.folder || '').toLowerCase();
            const score = keywords.reduce(function (acc, kw) {
                return acc + (name.includes(kw) ? 2 : 0) + (path.includes(kw) ? 1 : 0) + (folder.includes(kw) ? 1 : 0);
            }, 0);
            return { f: f, score: score };
        })
        .filter(function (x) { return x.score > 0; })
        .sort(function (a, b) { return b.score - a.score; })
        .slice(0, 15);

    return scored.map(function (x) {
        return {
            name: x.f.name,
            path: x.f.path,
            folder: x.f.folder,
            ext: x.f.ext,
            webUrl: x.f.webUrl
        };
    });
}

async function readFile(path, context) {
    const files = await loadIndex(context);
    if (!files) {
        return 'The search index is currently unavailable. Tell the user that file access is temporarily unavailable.';
    }

    const wanted = String(path || '');
    let entry = files.find(function (f) { return f.path === wanted; });
    if (!entry) {
        const lower = wanted.toLowerCase();
        entry = files.find(function (f) { return (f.path || '').toLowerCase() === lower; });
    }
    if (!entry) {
        return 'No file found at path: ' + wanted + '. Use search_files first and pass the exact "path" value it returns.';
    }

    if (isRestricted(entry.path)) {
        context.log('Blocked restricted file read:', entry.path);
        return 'That file is in a restricted folder containing sensitive personal/HR data and cannot be read.';
    }

    try {
        const graphToken = await getGraphToken();
        const siteId = await getSiteId(graphToken);
        const contentResponse = await fetch(
            'https://graph.microsoft.com/v1.0/sites/' + siteId + '/drive/items/' + entry.id + '/content',
            { headers: { 'Authorization': 'Bearer ' + graphToken } }
        );
        if (!contentResponse.ok) {
            return 'Could not download the file (status ' + contentResponse.status + ').';
        }

        const ext = (entry.ext || entry.name.split('.').pop() || '').toLowerCase();
        let textContent = '';

        if (ext === 'docx') {
            const mammoth = require('mammoth');
            const buffer = Buffer.from(await contentResponse.arrayBuffer());
            const result = await mammoth.extractRawText({ buffer: buffer });
            textContent = result.value;
        } else if (ext === 'xlsx' || ext === 'xls') {
            const XLSX = require('xlsx');
            const buffer = Buffer.from(await contentResponse.arrayBuffer());
            const workbook = XLSX.read(buffer, { type: 'buffer' });
            textContent = workbook.SheetNames.map(function (name) {
                return 'Sheet: ' + name + '\n' + XLSX.utils.sheet_to_csv(workbook.Sheets[name]);
            }).join('\n\n');
        } else if (ext === 'pdf') {
            const pdfParse = require('pdf-parse');
            const buffer = Buffer.from(await contentResponse.arrayBuffer());
            const data = await pdfParse(buffer);
            textContent = data.text;
        } else {
            textContent = await contentResponse.text();
        }

        context.log('Read file:', entry.path, 'type:', ext, 'length:', textContent.length);
        return '=== FILE: ' + entry.path + ' ===\n' + textContent.substring(0, 30000);
    } catch (e) {
        context.log('Error reading file:', entry.path, e.message);
        return 'Error reading the file: ' + e.message;
    }
}

async function runTool(name, input, context) {
    if (name === 'search_files') {
        const files = await loadIndex(context);
        if (!files) {
            return 'The search index is currently unavailable. Tell the user that file search is temporarily unavailable and to try again shortly.';
        }
        const hits = searchFiles(input && input.query, files);
        if (hits.length === 0) {
            return 'No files matched that query. Try different or broader keywords.';
        }
        return JSON.stringify(hits);
    }
    if (name === 'read_file') {
        return await readFile(input && input.path, context);
    }
    return 'Unknown tool: ' + name;
}

async function callClaude(baseBody, messages, tools) {
    const body = Object.assign({}, baseBody, { messages: messages });
    if (tools) {
        body.tools = tools;
    } else {
        delete body.tools;
    }
    const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'x-api-key': process.env.ANTHROPIC_API_KEY,
            'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify(body)
    });
    return await response.json();
}

module.exports = async function (context, req) {
    context.res = {
        headers: {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'POST, OPTIONS',
            'Access-Control-Allow-Headers': 'Content-Type, Authorization'
        }
    };

    if (req.method === 'OPTIONS') {
        context.res.status = 204;
        return;
    }

    try {
        const baseBody = req.body;
        const messages = (baseBody.messages || []).slice();

        let data;
        for (let i = 0; i < MAX_TOOL_ITERATIONS; i++) {
            data = await callClaude(baseBody, messages, TOOLS);

            if (data.stop_reason !== 'tool_use') {
                context.res.status = 200;
                context.res.body = JSON.stringify(data);
                return;
            }

            messages.push({ role: 'assistant', content: data.content });

            const toolResults = [];
            for (const block of data.content) {
                if (block.type === 'tool_use') {
                    context.log('Tool call:', block.name, JSON.stringify(block.input));
                    const result = await runTool(block.name, block.input, context);
                    toolResults.push({
                        type: 'tool_result',
                        tool_use_id: block.id,
                        content: result
                    });
                }
            }
            messages.push({ role: 'user', content: toolResults });
        }

        context.log('Tool-use loop hit iteration cap; forcing a final answer with no tools.');
        data = await callClaude(baseBody, messages, null);
        context.res.status = 200;
        context.res.body = JSON.stringify(data);
    } catch (err) {
        context.log('FATAL ERROR:', err.message);
        context.res.status = 500;
        context.res.body = JSON.stringify({ error: err.message });
    }
};

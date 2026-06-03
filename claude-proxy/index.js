// claude-proxy/index.js — Ask Standard Meat retrieval proxy
// Phase 1 (metadata-only): model-driven tool-use loop.
// Replaces the old full-library inventory injection (which overflowed the
// 200K prompt limit) — the assistant now calls search_files on demand instead
// of receiving every filename up front. This step adds search_files only;
// list_folder and read_file come in later steps.

// Content libs + Graph helpers are retained for read_file (added in a later step).
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');
const { getGraphToken, getSiteId } = require('../shared/graph');
// NOTE: @azure/storage-blob is required INSIDE loadIndex (not at module top) —
// per the build-index startup-crash lesson, heavy top-level requires can break
// the host<->worker gRPC handshake during the worker init window.

const MAX_TOOL_ITERATIONS = 5;

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
        description: 'Search the Standard Meat SharePoint document library by keyword. Matches against file names, paths, and folder names. Returns up to 15 matching files as metadata only: name, path, folder, file extension, and an openable SharePoint link. This tool does NOT return file contents. Use it to locate relevant files for the user\'s question. Reading file contents is not yet available, so base your answer on the file names and paths returned.',
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
    }
];

// Module-level cache (just a null reference at load time — no heavy work here).
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

function searchFiles(query, files) {
    const text = String(query || '').toLowerCase();
    let keywords = (text.match(/[a-z_]{2,}/g) || []).filter(w => !STOP_WORDS.has(w) && w.length >= 2);

    const expanded = new Set(keywords);
    for (const kw of keywords) {
        if (EXPANSIONS[kw]) EXPANSIONS[kw].forEach(e => expanded.add(e));
    }
    for (const phrase in EXPANSIONS) {
        if (phrase.includes(' ') && text.includes(phrase)) EXPANSIONS[phrase].forEach(e => expanded.add(e));
    }
    keywords = Array.from(expanded).filter(w => w.length >= 2);
    if (keywords.length === 0) return [];

    const scored = files
        .map(f => {
            const name = (f.name || '').toLowerCase();
            const path = (f.path || '').toLowerCase();
            const folder = (f.folder || '').toLowerCase();
            const score = keywords.reduce((acc, kw) =>
                acc + (name.includes(kw) ? 2 : 0) + (path.includes(kw) ? 1 : 0) + (folder.includes(kw) ? 1 : 0), 0);
            return { f, score };
        })
        .filter(x => x.score > 0)
        .sort((a, b) => b.score - a.score)
        .slice(0, 15);

    return scored.map(x => ({
        name: x.f.name,
        path: x.f.path,
        folder: x.f.folder,
        ext: x.f.ext,
        webUrl: x.f.webUrl
    }));
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
    return `Unknown tool: ${name}`;
}

async function callClaude(baseBody, messages, tools) {
    const body = { ...baseBody, messages };
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
        const messages = [...(baseBody.messages || [])];

        let data;
        for (let i = 0; i < MAX_TOOL_ITERATIONS; i++) {
            data = await callClaude(baseBody, messages, TOOLS);

            if (data.stop_reason !== 'tool_use') {
                context.res.status = 200;
                context.res.body = JSON.stringify(data);
                return;
            }

            // Claude wants a tool. Record its turn, run the tool(s), feed results back.
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

        // Iteration cap reached while still requesting tools — force a final text answer.
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

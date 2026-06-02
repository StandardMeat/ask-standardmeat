const mammoth = require('mammoth');
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');

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
        const userMessage = req.body.messages[req.body.messages.length - 1].content;
        context.log('User message:', userMessage);

        const tokenResponse = await fetch(
            `https://login.microsoftonline.com/${process.env.SHAREPOINT_TENANT_ID}/oauth2/v2.0/token`,
            {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams({
                    client_id: process.env.SHAREPOINT_CLIENT_ID,
                    client_secret: process.env.SHAREPOINT_CLIENT_SECRET,
                    scope: 'https://graph.microsoft.com/.default',
                    grant_type: 'client_credentials'
                })
            }
        );

        const tokenData = await tokenResponse.json();
        if (!tokenData.access_token) throw new Error('Failed to get Graph token');
        const graphToken = tokenData.access_token;

        const siteResponse = await fetch(
            'https://graph.microsoft.com/v1.0/sites/standardmeatco.sharepoint.com:/sites/ClaudePilot',
            { headers: { 'Authorization': `Bearer ${graphToken}` } }
        );
        const siteData = await siteResponse.json();
        if (!siteData.id) throw new Error('Failed to get site ID');
        const siteId = siteData.id;

        async function listAllFiles(folderPath = '') {
            if (folderPath.includes('.git') || folderPath.includes('node_modules')) return [];
            const url = folderPath
                ? `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeURIComponent(folderPath)}:/children?$top=500`
                : `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/children?$top=500`;
            const response = await fetch(url, { headers: { 'Authorization': `Bearer ${graphToken}` } });
            const data = await response.json();
            const items = data.value || [];
            let allFiles = [];
            for (const item of items) {
                if (item.file) {
                    allFiles.push({ ...item, folderPath });
                } else if (item.folder) {
                    if (item.name.startsWith('.')) continue;
                    const subPath = folderPath ? `${folderPath}/${item.name}` : item.name;
                    const subFiles = await listAllFiles(subPath);
                    allFiles = allFiles.concat(subFiles);
                }
            }
            return allFiles;
        }

        const allFiles = await listAllFiles();
        context.log('Total files found:', allFiles.length);

        const stopWords = new Set(['that','this','with','from','find','show','what','have','will','where','when','which','about','your','they','them','there','their','would','could','should','please','refer','look','tell','give','make','need','want','help','using','sends','pulling','scripts','script','file','files','process','actually','looking','supposed','then','another','how','many','are','in','on','at','of','an','the','do','does','did','was','were','been','being','can','also','any','all','some','our']);
        const userMessages = req.body.messages.filter(m => m.role === 'user');
        const latestMessage = userMessages.slice(-1).map(m => m.content).join(' ');
        let primaryKeywords = (latestMessage.toLowerCase().match(/[a-z_]{2,}/g) || []).filter(w => !stopWords.has(w) && w.length >= 2);

        // If the latest message has no substantive keywords (e.g. "yes", "ok", "do that one"),
        // fall back to the last 3 messages so we know what they're referring to.
        const searchText = primaryKeywords.length >= 2 ? latestMessage : userMessages.slice(-3).map(m => m.content).join(' ');
        let keywords = (searchText.toLowerCase().match(/[a-z_]{2,}/g) || []).filter(w => !stopWords.has(w) && w.length >= 2);
        context.log('Search source:', primaryKeywords.length >= 2 ? 'latest only' : 'last 3 messages');
        const expansions = {
            'accounts payable': ['ap'], 'accounts receivable': ['ar'], 'general ledger': ['gl'],
            'purchase order': ['po'], 'sales order': ['so'], 'inventory': ['inv'],
            'payable': ['ap'], 'receivable': ['ar'], 'vendor': ['vend'],
            'ap': ['payable','accounts'], 'ar': ['receivable','accounts'], 'gl': ['ledger'],
            'po': ['purchase'], 'edi': ['edi']
        };
        const expandedKeywords = new Set(keywords);
        for (const kw of keywords) if (expansions[kw]) expansions[kw].forEach(e => expandedKeywords.add(e));
        const lowerText = searchText.toLowerCase();
        for (const phrase in expansions) {
            if (phrase.includes(' ') && lowerText.includes(phrase)) expansions[phrase].forEach(e => expandedKeywords.add(e));
        }
        keywords = Array.from(expandedKeywords).filter(w => w.length >= 2);
        context.log('Keywords:', keywords.join(','));

        const scoredFiles = allFiles.map(f => {
            const name = f.name.toLowerCase();
            const path = (f.folderPath || '').toLowerCase();
            const score = keywords.reduce((acc, kw) => acc + (name.includes(kw) ? 2 : 0) + (path.includes(kw) ? 1 : 0), 0);
            return { ...f, score };
        }).filter(f => f.score > 0).sort((a, b) => b.score - a.score);
        context.log('Matching files:', scoredFiles.length);

        // Always build complete file inventory grouped by folder
        const filesByFolder = {};
        for (const f of allFiles) {
            const folder = f.folderPath || 'root';
            if (!filesByFolder[folder]) filesByFolder[folder] = [];
            filesByFolder[folder].push(f.name);
        }
        let fileInventory = '\n\n=== COMPLETE FILE INVENTORY (all files in SharePoint library) ===\n';
        for (const folder in filesByFolder) {
            fileInventory += `\n${folder}/:\n  ${filesByFolder[folder].join(', ')}`;
        }
        context.log('File inventory built:', allFiles.length, 'files');

        let fileContents = '';
        const topFiles = scoredFiles.slice(0, 3);
        context.log('Files to read:', topFiles.map(f => `${f.folderPath}/${f.name} (score:${f.score})`).join(', '));

        for (const file of topFiles) {
            try {
                const ext = file.name.toLowerCase().split('.').pop();
                const contentResponse = await fetch(
                    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${file.id}/content`,
                    { headers: { 'Authorization': `Bearer ${graphToken}` } }
                );

                let textContent = '';

                if (ext === 'docx') {
                    const buffer = Buffer.from(await contentResponse.arrayBuffer());
                    const result = await mammoth.extractRawText({ buffer });
                    textContent = result.value;
                } else if (ext === 'xlsx' || ext === 'xls') {
                    const buffer = Buffer.from(await contentResponse.arrayBuffer());
                    const workbook = XLSX.read(buffer, { type: 'buffer' });
                    textContent = workbook.SheetNames.map(name => {
                        return `Sheet: ${name}\n${XLSX.utils.sheet_to_csv(workbook.Sheets[name])}`;
                    }).join('\n\n');
                } else if (ext === 'pdf') {
                    const buffer = Buffer.from(await contentResponse.arrayBuffer());
                    const data = await pdfParse(buffer);
                    textContent = data.text;
                } else {
                    textContent = await contentResponse.text();
                }

                fileContents += `\n\n=== FILE: ${file.folderPath}/${file.name} ===\n${textContent.substring(0, 30000)}`;
                context.log('Read file:', file.name, 'type:', ext, 'length:', textContent.length);
            } catch (e) {
                context.log('Error reading:', file.name, e.message);
            }
        }

        const enhancedMessages = [...req.body.messages];
        let contextBlock = `Context from SharePoint:${fileInventory}\n\n`;
        if (fileContents) {
            contextBlock += `Detailed contents of files most relevant to your question:${fileContents}\n\n`;
        }
        contextBlock += `---\n\n`;
        enhancedMessages[enhancedMessages.length - 1] = {
            role: 'user',
            content: `${contextBlock}User question: ${userMessage}`
        };

        console.error('SZ history(messages)=', JSON.stringify(req.body.messages).length);
        console.error('SZ allFiles count=', (typeof allFiles !== 'undefined' ? allFiles.length : 'n/a'));
        console.error('SZ fileInventory=', fileInventory.length);
        console.error('SZ topFiles count=', topFiles.length);
        console.error('SZ fileContents=', fileContents.length);
        console.error('SZ contextBlock=', contextBlock.length);
        console.error('SZ enhancedBody=', JSON.stringify({ ...req.body, messages: enhancedMessages }).length);

        const claudeBody = { ...req.body, messages: enhancedMessages };
        const response = await fetch('https://api.anthropic.com/v1/messages', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'x-api-key': process.env.ANTHROPIC_API_KEY,
                'anthropic-version': '2023-06-01'
            },
            body: JSON.stringify(claudeBody)
        });

        const data = await response.json();
        context.res.status = 200;
        context.res.body = JSON.stringify(data);
    } catch (err) {
        context.log('FATAL ERROR:', err.message);
        context.res.status = 500;
        context.res.body = JSON.stringify({ error: err.message });
    }
};

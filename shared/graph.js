async function getGraphToken() {
    const response = await fetch(
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
    const data = await response.json();
    if (!data.access_token) throw new Error('Failed to get Graph token');
    return data.access_token;
}

async function getSiteId(token) {
    const response = await fetch(
        'https://graph.microsoft.com/v1.0/sites/standardmeatco.sharepoint.com:/sites/ClaudePilot',
        { headers: { 'Authorization': `Bearer ${token}` } }
    );
    const data = await response.json();
    if (!data.id) throw new Error('Failed to get site ID');
    return data.id;
}

async function listAllFiles(token, siteId, folderPath = '') {
    if (folderPath.includes('.git') || folderPath.includes('node_modules')) return [];
    let url = folderPath
        ? `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodeURIComponent(folderPath)}:/children?$top=500`
        : `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/children?$top=500`;
    let allFiles = [];
    while (url) {
        const response = await fetch(url, { headers: { 'Authorization': `Bearer ${token}` } });
        const data = await response.json();
        const items = data.value || [];
        for (const item of items) {
            if (item.file) {
                allFiles.push({ ...item, folderPath });
            } else if (item.folder) {
                if (item.name.startsWith('.')) continue;
                const subPath = folderPath ? `${folderPath}/${item.name}` : item.name;
                const subFiles = await listAllFiles(token, siteId, subPath);
                allFiles = allFiles.concat(subFiles);
            }
        }
        url = data['@odata.nextLink'] || null;
    }
    return allFiles;
}

module.exports = { getGraphToken, getSiteId, listAllFiles };

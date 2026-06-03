const { BlobServiceClient } = require('@azure/storage-blob');
const { getGraphToken, getSiteId, listAllFiles } = require('../shared/graph');

const SOURCE = 'ClaudePilot Workspace';
const CONTAINER = 'index';
const BLOB_NAME = 'file-index.json';

module.exports = async function (context, myTimer) {
    try {
        const token = await getGraphToken();
        const siteId = await getSiteId(token);
        const rawFiles = await listAllFiles(token, siteId);
        context.log('Enumerated files:', rawFiles.length);

        const files = rawFiles.map(f => {
            const folder = f.folderPath || '';
            const path = folder ? `${folder}/${f.name}` : f.name;
            const dot = f.name.lastIndexOf('.');
            const ext = dot > 0 ? f.name.slice(dot + 1).toLowerCase() : '';
            return {
                id: f.id,
                name: f.name,
                path,
                folder,
                ext,
                size: f.size,
                modified: f.lastModifiedDateTime,
                webUrl: f.webUrl,
                source: SOURCE
            };
        });

        const output = {
            generatedAt: new Date().toISOString(),
            count: files.length,
            files
        };
        const body = JSON.stringify(output);

        const blobService = BlobServiceClient.fromConnectionString(process.env.AzureWebJobsStorage);
        const container = blobService.getContainerClient(CONTAINER);
        await container.createIfNotExists();
        const blob = container.getBlockBlobClient(BLOB_NAME);
        await blob.upload(body, Buffer.byteLength(body), {
            blobHTTPHeaders: { blobContentType: 'application/json' }
        });

        context.log(`build-index done: count=${output.count} generatedAt=${output.generatedAt} blob=${CONTAINER}/${BLOB_NAME}`);
    } catch (err) {
        context.log('build-index error:', err.message);
        throw err;
    }
};

const uploadSection = document.getElementById('upload-section');
const editorSection = document.getElementById('editor-section');
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const browseBtn = document.getElementById('browse-btn');
const fileNameDisplay = document.getElementById('file-name-display');
const cancelBtn = document.getElementById('cancel-btn');
const editSaveBtn = document.getElementById('edit-save-btn');
const metadataContainer = document.getElementById('metadata-container');
const loadingOverlay = document.getElementById('loading-overlay');
const loadingText = document.getElementById('loading-text');

let currentFile = null;
let currentZip = null;
let isEditMode = false;

const metadataConfig = [
    { id: 'dc:title', label: 'Title', file: 'docProps/core.xml', ns: 'http://purl.org/dc/elements/1.1/', isCoreNode: true },
    { id: 'dc:subject', label: 'Subject', file: 'docProps/core.xml', ns: 'http://purl.org/dc/elements/1.1/', isCoreNode: true },
    { id: 'dc:creator', label: 'Creator (Author)', file: 'docProps/core.xml', ns: 'http://purl.org/dc/elements/1.1/', isCoreNode: true },
    { id: 'dc:description', label: 'Description', file: 'docProps/core.xml', ns: 'http://purl.org/dc/elements/1.1/', isCoreNode: true },
    { id: 'cp:keywords', label: 'Keywords', file: 'docProps/core.xml', ns: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', isCoreNode: true },
    { id: 'cp:lastModifiedBy', label: 'Last Modified By', file: 'docProps/core.xml', ns: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', isCoreNode: true },
    { id: 'cp:category', label: 'Category', file: 'docProps/core.xml', ns: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', isCoreNode: true },
    { id: 'cp:contentStatus', label: 'Status', file: 'docProps/core.xml', ns: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', isCoreNode: true },
    { id: 'dcterms:created', label: 'Creation Date (ISO 8601)', file: 'docProps/core.xml', ns: 'http://purl.org/dc/terms/', isCoreNode: true },
    { id: 'dcterms:modified', label: 'Last Modified Date (ISO 8601)', file: 'docProps/core.xml', ns: 'http://purl.org/dc/terms/', isCoreNode: true },
    { id: 'Company', label: 'Company', file: 'docProps/app.xml', isAppNode: true },
    { id: 'Manager', label: 'Manager', file: 'docProps/app.xml', isAppNode: true },
    { id: 'Application', label: 'Application', file: 'docProps/app.xml', isAppNode: true },
    { id: 'TotalTime', label: 'Summary Editing Time (min)', file: 'docProps/app.xml', isAppNode: true }
];

const extractedData = new Map();
const originalData = new Map();

const NAMESPACES = {
    dc: 'http://purl.org/dc/elements/1.1/',
    dcterms: 'http://purl.org/dc/terms/',
    cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    ep: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    xsi: 'http://www.w3.org/2001/XMLSchema-instance'
};

document.addEventListener('DOMContentLoaded', () => {
    browseBtn.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            const file = e.dataTransfer.files[0];
            if (file.name.endsWith('.docx') || file.name.endsWith('.docm')) {
                handleFile(file);
            } else {
                alert('Unsupported file format. Please upload a .docx or .docm file.');
            }
        }
    });

    cancelBtn.addEventListener('click', () => {
        resetApp();
    });

    editSaveBtn.addEventListener('click', async () => {
        if (isEditMode) {
            await saveAndDownload();
        } else {
            enterEditMode();
        }
    });
});

async function handleFile(file) {
    currentFile = file;
    fileNameDisplay.textContent = file.name;

    uploadSection.classList.remove('active');
    uploadSection.classList.add('hidden');
    editorSection.classList.remove('hidden');
    editorSection.classList.add('active');

    isEditMode = false;
    editSaveBtn.textContent = 'Edit';
    editSaveBtn.classList.remove('btn-success');
    editSaveBtn.classList.add('btn-primary');

    await parseMetadata();
}

function showLoading(text) {
    loadingText.textContent = text;
    loadingOverlay.classList.remove('hidden');
}

function hideLoading() {
    loadingOverlay.classList.add('hidden');
}

function resetApp() {
    currentFile = null;
    currentZip = null;
    isEditMode = false;
    extractedData.clear();
    fileInput.value = '';

    editorSection.classList.remove('active');
    setTimeout(() => {
        editorSection.classList.add('hidden');
        uploadSection.classList.remove('hidden');
        uploadSection.classList.add('active');
    }, 300);
}

let xmlDocs = {};

async function parseMetadata() {
    showLoading('Parsing file...');
    try {
        const fileData = await currentFile.arrayBuffer();
        currentZip = await JSZip.loadAsync(fileData);

        extractedData.clear();
        originalData.clear();
        xmlDocs = {};

        const getXmlDoc = async (path) => {
            if (currentZip.file(path)) {
                const xmlString = await currentZip.file(path).async('string');
                const parser = new DOMParser();
                return parser.parseFromString(xmlString, 'text/xml');
            }
            return null;
        };

        xmlDocs['docProps/core.xml'] = await getXmlDoc('docProps/core.xml');
        xmlDocs['docProps/app.xml'] = await getXmlDoc('docProps/app.xml');

        metadataConfig.forEach(config => {
            let value = '';
            const doc = xmlDocs[config.file];
            if (doc) {
                let node;

                if (config.isCoreNode) {
                    const localName = config.id.split(':')[1] || config.id;
                    const elList = config.ns ? doc.getElementsByTagNameNS(config.ns, localName) : doc.getElementsByTagName(config.id);
                    if (elList && elList.length > 0) {
                        node = elList[0];
                    } else if (doc.getElementsByTagName(config.id).length > 0) {
                        node = doc.getElementsByTagName(config.id)[0];
                    }
                } else if (config.isAppNode) {
                    const elList = doc.getElementsByTagName(config.id);
                    if (elList && elList.length > 0) {
                        node = elList[0];
                    }
                }

                if (node) {
                    value = node.textContent || '';
                }
            }
            extractedData.set(config.id, value);
            originalData.set(config.id, value);
        });

        renderMetadataGrid();

    } catch (err) {
        console.error("Error reading file:", err);
        alert('Error parsing the file. It might be corrupted or an invalid Word document.');
        resetApp();
    } finally {
        hideLoading();
    }
}

function renderMetadataGrid() {
    metadataContainer.innerHTML = '';

    metadataConfig.forEach(config => {
        const value = extractedData.get(config.id) || '';

        const itemDiv = document.createElement('div');
        itemDiv.className = 'metadata-item';

        const label = document.createElement('div');
        label.className = 'metadata-label';
        label.textContent = config.label;

        if (isEditMode) {
            const input = document.createElement('input');
            input.type = 'text';
            input.className = 'metadata-input';
            input.value = value;
            input.dataset.id = config.id;
            input.addEventListener('change', (e) => {
                extractedData.set(config.id, e.target.value.trim());
            });
            itemDiv.appendChild(label);
            itemDiv.appendChild(input);
        } else {
            const valueDiv = document.createElement('div');
            valueDiv.className = 'metadata-value';
            valueDiv.textContent = value || '—';
            itemDiv.appendChild(label);
            itemDiv.appendChild(valueDiv);
        }

        metadataContainer.appendChild(itemDiv);
    });
}

function enterEditMode() {
    isEditMode = true;
    editSaveBtn.textContent = 'Save';
    editSaveBtn.style.backgroundColor = 'var(--success-color)';
    editSaveBtn.style.boxShadow = '0 4px 14px 0 rgba(16, 185, 129, 0.3)';
    renderMetadataGrid();
}

function createMissingNode(doc, config) {
    if (config.isCoreNode) {
        const parts = config.id.split(':');
        const prefix = parts[0];
        const localName = parts[1];
        const nsURI = NAMESPACES[prefix];

        if (nsURI) {
            return doc.createElementNS(nsURI, config.id);
        } else {
            return doc.createElement(config.id);
        }
    } else {
        const nsURI = NAMESPACES['ep'];
        if (nsURI) {
            return doc.createElementNS(nsURI, config.id);
        }
        return doc.createElement(config.id);
    }
}

async function saveAndDownload() {
    showLoading('Updating file...');

    try {
        let hasChanges = false;

        const inputs = document.querySelectorAll('.metadata-input');
        inputs.forEach(input => {
            const currentVal = input.value.trim();
            const originalVal = originalData.get(input.dataset.id) || '';

            extractedData.set(input.dataset.id, currentVal);

            if (currentVal !== originalVal) {
                hasChanges = true;
            }
        });

        if (!hasChanges) {
            isEditMode = false;
            editSaveBtn.textContent = 'Edit';
            editSaveBtn.style.backgroundColor = 'var(--accent-color)';
            editSaveBtn.style.boxShadow = '0 4px 14px 0 var(--accent-glow)';
            renderMetadataGrid();
            hideLoading();
            return;
        }

        metadataConfig.forEach(config => {
            const doc = xmlDocs[config.file];
            if (!doc) return;

            const newValue = extractedData.get(config.id) || '';
            let node = null;

            if (config.isCoreNode) {
                const localName = config.id.split(':')[1] || config.id;
                const elList = config.ns ? doc.getElementsByTagNameNS(config.ns, localName) : doc.getElementsByTagName(config.id);
                if (elList && elList.length > 0) {
                    node = elList[0];
                } else if (doc.getElementsByTagName(config.id).length > 0) {
                    node = doc.getElementsByTagName(config.id)[0];
                }
            } else if (config.isAppNode) {
                const elList = doc.getElementsByTagName(config.id);
                if (elList && elList.length > 0) {
                    node = elList[0];
                }
            }

            if (node) {
                node.textContent = newValue;
            } else if (newValue !== '') {
                const rootNode = doc.documentElement;
                if (rootNode) {
                    const newNode = createMissingNode(doc, config);
                    newNode.textContent = newValue;
                    rootNode.appendChild(newNode);
                }
            }
        });

        const serializer = new XMLSerializer();
        for (const [path, doc] of Object.entries(xmlDocs)) {
            if (doc) {
                let xmlString = serializer.serializeToString(doc);
                if (!xmlString.startsWith('<?xml')) {
                    xmlString = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + xmlString;
                }
                currentZip.file(path, xmlString);
            }
        }

        const newContent = await currentZip.generateAsync({
            type: "blob",
            compression: "DEFLATE",
            compressionOptions: {
                level: 6
            }
        });

        const newFileName = "edited_" + (currentFile.name || "document.docx");

        const finalBlob = new Blob([newContent], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
        saveAs(finalBlob, newFileName);

        isEditMode = false;
        editSaveBtn.textContent = 'Edit';
        editSaveBtn.style.backgroundColor = 'var(--accent-color)';
        editSaveBtn.style.boxShadow = '0 4px 14px 0 var(--accent-glow)';
        renderMetadataGrid();

    } catch (err) {
        console.error("Error saving file:", err);
        alert('An error occurred while saving the file. See console for details.');
    } finally {
        hideLoading();
    }
}

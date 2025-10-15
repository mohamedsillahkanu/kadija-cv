const GOOGLE_APPS_SCRIPT_PROXY = 'https://script.google.com/macros/s/AKfycbyRXc68PYSmDSCYPakAS8Gc9R00kRjmsMYoDujPYMItyquXfwMAJUtRqXLiDlNWbSfDQQ/exec';
const state = { config: null, isLoggedIn: false };
let datasetsData = [];
let selectedDataElements = new Set();
let orgUnitsData = [];
let selectedOrgUnits = new Set();

function makeProxyRequest(targetUrl, options = {}) {
    let proxyUrl = `${GOOGLE_APPS_SCRIPT_PROXY}?url=${encodeURIComponent(targetUrl)}`;
    if (options.headers && options.headers.Authorization) {
        proxyUrl += `&Authorization=${encodeURIComponent(options.headers.Authorization)}`;
    }
    return fetch(proxyUrl, { method: options.method || 'GET', body: options.body });
}

async function fetchWithAuth(url, options = {}) {
    if (!state.config) throw new Error('Not authenticated');
    const authHeader = 'Basic ' + btoa(`${state.config.username}:${state.config.password}`);
    return makeProxyRequest(url, { ...options, headers: { 'Authorization': authHeader, 'Content-Type': 'application/json', ...options.headers } });
}

function showNotification(message, type) {
    const notification = document.getElementById('notification');
    if (!notification) return;
    notification.textContent = message;
    notification.className = `notification ${type} show`;
    setTimeout(() => notification.classList.remove('show'), 4000);
}

function setProgress(percent) {
    const progressFill = document.getElementById('progressFill');
    if (progressFill) progressFill.style.width = percent + '%';
}

function showProgress() {
    const progressSection = document.getElementById('progressSection');
    if (progressSection) progressSection.classList.add('show');
}

function log(message, type = 'info') {
    const container = document.getElementById('logContainer');
    if (!container) return;
    const entry = document.createElement('div');
    entry.className = `log-entry ${type}`;
    entry.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
    container.appendChild(entry);
    container.scrollTop = container.scrollHeight;
}

function clearLogs() {
    const container = document.getElementById('logContainer');
    if (container) container.innerHTML = '';
}

function selectPeriodType(type) {
    document.querySelectorAll('.period-tab').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.period-inputs').forEach(input => input.classList.remove('active'));
    event.target.classList.add('active');
    const targetInput = document.getElementById(`${type}-inputs`);
    if (targetInput) targetInput.classList.add('active');
}

async function loadDatasetsWithElements() {
    try {
        log('📋 Fetching datasets from DHIS2...', 'info');
        showNotification('Loading datasets...', 'info');
        const url = `${state.config.instanceUrl}/api/dataSets.json?fields=id,displayName,name,periodType,dataSetElements[dataElement[id,displayName,name]]&paging=false`;
        const response = await fetchWithAuth(url);
        if (!response.ok) throw new Error(`Failed to fetch datasets: ${response.status}`);
        const data = await response.json();
        if (!data.dataSets || !Array.isArray(data.dataSets)) {
            showNotification('No datasets found in response', 'error');
            return;
        }
        datasetsData = data.dataSets;
        renderAvailableItems(datasetsData);
        renderSelectedItems();
        document.getElementById('availableSearch').style.display = 'block';
        document.getElementById('addAllBtn').style.display = 'block';
        log(`✅ Loaded ${data.dataSets.length} datasets`, 'success');
        showNotification(`Loaded ${data.dataSets.length} datasets`, 'success');
    } catch (error) {
        log(`❌ Error: ${error.message}`, 'error');
        showNotification(`Failed to load datasets: ${error.message}`, 'error');
    }
}

function renderAvailableItems(datasets) {
    const container = document.getElementById('availablePanel');
    let totalElements = 0;
    datasets.forEach(ds => {
        if (ds.dataSetElements) {
            ds.dataSetElements.forEach(dse => {
                if (dse.dataElement && !selectedDataElements.has(dse.dataElement.id)) totalElements++;
            });
        }
    });
    if (totalElements === 0) {
        container.innerHTML = '<div class="empty-state"><div class="empty-state-icon">✨</div>All items have been selected!</div>';
        document.getElementById('availableCount').textContent = '0 items';
        return;
    }
    container.innerHTML = '';
    datasets.forEach(ds => {
        if (!ds.dataSetElements || ds.dataSetElements.length === 0) return;
        const availableElements = ds.dataSetElements.filter(dse => dse.dataElement && !selectedDataElements.has(dse.dataElement.id));
        if (availableElements.length === 0) return;
        const datasetDiv = document.createElement('div');
        datasetDiv.className = 'dataset-item';
        datasetDiv.dataset.id = ds.id;
        const toggle = document.createElement('span');
        toggle.className = 'item-toggle';
        toggle.textContent = '▶';
        toggle.onclick = (e) => { e.stopPropagation(); toggleDatasetExpand(ds.id); };
        const nameSpan = document.createElement('span');
        nameSpan.className = 'item-name';
        const periodType = ds.periodType ? ` (${ds.periodType})` : '';
        nameSpan.textContent = `${ds.displayName || ds.name || ds.id}${periodType} - ${availableElements.length} elements`;
        datasetDiv.appendChild(toggle);
        datasetDiv.appendChild(nameSpan);
        datasetDiv.onclick = (e) => { if (e.target.className !== 'item-toggle') toggleDatasetExpand(ds.id); };
        container.appendChild(datasetDiv);
        const elementsContainer = document.createElement('div');
        elementsContainer.className = 'data-elements-container';
        elementsContainer.id = `elements-${ds.id}`;
        availableElements.forEach(dse => {
            if (dse.dataElement) {
                const elementDiv = document.createElement('div');
                elementDiv.className = 'data-element-item';
                elementDiv.dataset.id = dse.dataElement.id;
                elementDiv.dataset.name = (dse.dataElement.displayName || dse.dataElement.name || '').toLowerCase();
                const elNameSpan = document.createElement('span');
                elNameSpan.className = 'item-name';
                elNameSpan.textContent = dse.dataElement.displayName || dse.dataElement.name || dse.dataElement.id;
                const addBtn = document.createElement('button');
                addBtn.className = 'add-btn';
                addBtn.innerHTML = '➕ Add';
                addBtn.onclick = (e) => { e.stopPropagation(); addItem(dse.dataElement.id); };
                elementDiv.onclick = () => addItem(dse.dataElement.id);
                elementDiv.appendChild(elNameSpan);
                elementDiv.appendChild(addBtn);
                elementsContainer.appendChild(elementDiv);
            }
        });
        container.appendChild(elementsContainer);
    });
    document.getElementById('availableCount').textContent = `${totalElements} items`;
}

function toggleDatasetExpand(datasetId) {
    const elementsContainer = document.getElementById(`elements-${datasetId}`);
    const toggle = document.querySelector(`.dataset-item[data-id="${datasetId}"] .item-toggle`);
    if (elementsContainer && toggle) {
        if (elementsContainer.classList.contains('expanded')) {
            elementsContainer.classList.remove('expanded');
            toggle.textContent = '▶';
        } else {
            elementsContainer.classList.add('expanded');
            toggle.textContent = '▼';
        }
    }
}

function renderSelectedItems() {
    const container = document.getElementById('selectedPanel');
    const selectedItems = [];
    datasetsData.forEach(ds => {
        if (ds.dataSetElements) {
            ds.dataSetElements.forEach(dse => {
                if (dse.dataElement && selectedDataElements.has(dse.dataElement.id)) {
                    selectedItems.push({ id: dse.dataElement.id, name: dse.dataElement.displayName || dse.dataElement.name || dse.dataElement.id, datasetName: ds.displayName || ds.name || ds.id });
                }
            });
        }
    });
    if (selectedItems.length === 0) {
        container.innerHTML = '';
        document.getElementById('selectedCount').textContent = '0 items';
        return;
    }
    container.innerHTML = '';
    selectedItems.forEach(item => {
        const itemDiv = document.createElement('div');
        itemDiv.className = 'selected-item';
        itemDiv.dataset.id = item.id;
        itemDiv.dataset.name = item.name.toLowerCase();
        const nameSpan = document.createElement('span');
        nameSpan.className = 'item-name';
        nameSpan.textContent = item.name;
        nameSpan.title = `From: ${item.datasetName}`;
        const removeBtn = document.createElement('button');
        removeBtn.className = 'remove-btn';
        removeBtn.innerHTML = '❌';
        removeBtn.onclick = (e) => { e.stopPropagation(); removeItem(item.id); };
        itemDiv.appendChild(nameSpan);
        itemDiv.appendChild(removeBtn);
        container.appendChild(itemDiv);
    });
    document.getElementById('selectedCount').textContent = `${selectedItems.length} items`;
}

function addItem(itemId) {
    selectedDataElements.add(itemId);
    renderAvailableItems(datasetsData);
    renderSelectedItems();
}

function removeItem(itemId) {
    selectedDataElements.delete(itemId);
    renderAvailableItems(datasetsData);
    renderSelectedItems();
}

function addAllItems() {
    datasetsData.forEach(ds => { if (ds.dataSetElements) { ds.dataSetElements.forEach(dse => { if (dse.dataElement) selectedDataElements.add(dse.dataElement.id); }); } });
    renderAvailableItems(datasetsData);
    renderSelectedItems();
}

function clearAllItems() {
    selectedDataElements.clear();
    renderAvailableItems(datasetsData);
    renderSelectedItems();
}

function filterAvailableItems() {
    const searchTerm = document.getElementById('availableSearch').value.toLowerCase();
    const datasets = document.querySelectorAll('#availablePanel .dataset-item');
    datasets.forEach(datasetDiv => {
        const datasetId = datasetDiv.dataset.id;
        const elementsContainer = document.getElementById(`elements-${datasetId}`);
        if (!elementsContainer) return;
        const elements = elementsContainer.querySelectorAll('.data-element-item');
        let hasVisibleElements = false;
        elements.forEach(element => {
            const name = element.dataset.name;
            if (name && name.includes(searchTerm)) {
                element.style.display = 'flex';
                hasVisibleElements = true;
            } else {
                element.style.display = 'none';
            }
        });
        const datasetName = datasetDiv.querySelector('.item-name').textContent.toLowerCase();
        if (datasetName.includes(searchTerm) || hasVisibleElements) {
            datasetDiv.style.display = 'flex';
            if (hasVisibleElements && searchTerm) {
                elementsContainer.classList.add('expanded');
                const toggle = datasetDiv.querySelector('.item-toggle');
                if (toggle) toggle.textContent = '▼';
            }
        } else {
            datasetDiv.style.display = 'none';
        }
    });
}

function filterSelectedItems() {
    const searchTerm = document.getElementById('selectedSearch').value.toLowerCase();
    document.querySelectorAll('#selectedPanel .selected-item').forEach(item => {
        const name = item.dataset.name;
        item.style.display = name.includes(searchTerm) ? 'flex' : 'none';
    });
}

async function loadOrgUnitsTree() {
    try {
        log('🏢 Fetching organisation units...', 'info');
        const url = `${state.config.instanceUrl}/api/organisationUnits.json?fields=id,displayName,level,parent[id],children[id]&paging=false`;
        const response = await fetchWithAuth(url);
        if (!response.ok) throw new Error('Failed to fetch organisation units');
        const data = await response.json();
        orgUnitsData = data.organisationUnits;
        const tree = buildOrgUnitTree(orgUnitsData);
        renderOrgUnitTree(tree);
        log(`✅ Loaded ${orgUnitsData.length} organisation units`, 'success');
        showNotification(`Loaded ${orgUnitsData.length} organisation units`, 'success');
    } catch (error) {
        log(`❌ Error: ${error.message}`, 'error');
        showNotification('Failed to load organisation units', 'error');
    }
}

function buildOrgUnitTree(orgUnits) {
    const orgUnitMap = {};
    const rootUnits = [];
    orgUnits.forEach(ou => { orgUnitMap[ou.id] = { ...ou, children: [] }; });
    orgUnits.forEach(ou => {
        if (ou.parent && orgUnitMap[ou.parent.id]) {
            orgUnitMap[ou.parent.id].children.push(orgUnitMap[ou.id]);
        } else {
            rootUnits.push(orgUnitMap[ou.id]);
        }
    });
    return rootUnits;
}

function renderOrgUnitTree(tree) {
    const container = document.getElementById('orgUnitTree');
    container.innerHTML = '';
    tree.forEach(ou => container.appendChild(createOrgUnitNode(ou)));
}

function createOrgUnitNode(orgUnit) {
    const div = document.createElement('div');
    const itemDiv = document.createElement('div');
    itemDiv.className = `org-unit-item`;
    itemDiv.dataset.id = orgUnit.id;
    const toggle = document.createElement('span');
    toggle.className = 'org-unit-toggle';
    if (orgUnit.children && orgUnit.children.length > 0) {
        toggle.textContent = '▶';
        toggle.onclick = (e) => { e.stopPropagation(); toggleOrgUnit(orgUnit.id); };
    } else {
        toggle.textContent = '  ';
        toggle.style.cursor = 'default';
    }
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.className = 'org-unit-checkbox';
    checkbox.dataset.id = orgUnit.id;
    checkbox.checked = selectedOrgUnits.has(orgUnit.id);
    checkbox.onchange = (e) => { e.stopPropagation(); toggleOrgUnitSelection(orgUnit.id, e.target.checked); };
    const label = document.createElement('span');
    label.className = 'org-unit-label';
    label.textContent = orgUnit.displayName;
    label.onclick = () => { checkbox.checked = !checkbox.checked; toggleOrgUnitSelection(orgUnit.id, checkbox.checked); };
    itemDiv.appendChild(toggle);
    itemDiv.appendChild(checkbox);
    itemDiv.appendChild(label);
    div.appendChild(itemDiv);
    if (orgUnit.children && orgUnit.children.length > 0) {
        const childrenDiv = document.createElement('div');
        childrenDiv.className = 'org-unit-children';
        childrenDiv.id = `children-${orgUnit.id}`;
        orgUnit.children.forEach(child => { childrenDiv.appendChild(createOrgUnitNode(child)); });
        div.appendChild(childrenDiv);
    }
    return div;
}

function toggleOrgUnit(id) {
    const childrenDiv = document.getElementById(`children-${id}`);
    const toggle = document.querySelector(`[data-id="${id}"] .org-unit-toggle`);
    if (childrenDiv.classList.contains('expanded')) {
        childrenDiv.classList.remove('expanded');
        toggle.textContent = '▶';
    } else {
        childrenDiv.classList.add('expanded');
        toggle.textContent = '▼';
    }
}

function toggleOrgUnitSelection(id, selected) {
    if (selected) { selectedOrgUnits.add(id); } else { selectedOrgUnits.delete(id); }
    const item = document.querySelector(`.org-unit-item[data-id="${id}"]`);
    if (item) {
        if (selected) { item.classList.add('selected'); } else { item.classList.remove('selected'); }
    }
}

function expandAllOrgUnits() {
    document.querySelectorAll('.org-unit-children').forEach(div => div.classList.add('expanded'));
    document.querySelectorAll('.org-unit-toggle').forEach(toggle => { if (toggle.textContent === '▶') toggle.textContent = '▼'; });
}

function collapseAllOrgUnits() {
    document.querySelectorAll('.org-unit-children').forEach(div => div.classList.remove('expanded'));
    document.querySelectorAll('.org-unit-toggle').forEach(toggle => { if (toggle.textContent === '▼') toggle.textContent = '▶'; });
}

function selectAllOrgUnits() {
    document.querySelectorAll('.org-unit-checkbox').forEach(checkbox => { checkbox.checked = true; selectedOrgUnits.add(checkbox.dataset.id); });
    document.querySelectorAll('.org-unit-item').forEach(item => item.classList.add('selected'));
}

function clearAllOrgUnits() {
    document.querySelectorAll('.org-unit-checkbox').forEach(checkbox => checkbox.checked = false);
    selectedOrgUnits.clear();
    document.querySelectorAll('.org-unit-item').forEach(item => item.classList.remove('selected'));
}

function getPeriodFromForm() {
    const activeTab = document.querySelector('.period-tab.active');
    if (!activeTab) return '';
    const tabText = activeTab.textContent.trim().toLowerCase();
    if (tabText === 'monthly') {
        const periodInput = document.querySelector('#monthly-inputs input[name="period"]');
        const period = periodInput?.value;
        if (period) return `period=${period.replace('-', '')}`;
    } else if (tabText === 'date range') {
        const startInput = document.querySelector('#range-inputs input[name="periodStart"]');
        const endInput = document.querySelector('#range-inputs input[name="periodEnd"]');
        const start = startInput?.value;
        const end = endInput?.value;
        if (start && end) return `startDate=${start}&endDate=${end}`;
    } else if (tabText === 'quarterly') {
        const yearInput = document.querySelector('#quarterly-inputs input[name="year"]');
        const quarterSelect = document.querySelector('#quarterly-inputs select[name="quarter"]');
        const year = yearInput?.value;
        const quarter = quarterSelect?.value;
        if (year && quarter) return `period=${year}${quarter}`;
    } else if (tabText === 'yearly') {
        const startYearInput = document.querySelector('#yearly-inputs input[name="startYear"]');
        const endYearInput = document.querySelector('#yearly-inputs input[name="endYear"]');
        const startYear = startYearInput?.value;
        const endYear = endYearInput?.value;
        if (startYear && endYear) return `startDate=${startYear}-01-01&endDate=${endYear}-12-31`;
        else if (startYear) return `period=${startYear}`;
    }
    return '';
}

async function executeDownloadDatasets() {
    const selectedDataElementIds = Array.from(selectedDataElements);
    const selectedOrgUnitIds = Array.from(selectedOrgUnits);
    if (selectedDataElementIds.length === 0) { showNotification('Please select at least one data element', 'error'); return; }
    if (selectedOrgUnitIds.length === 0) { showNotification('Please select at least one organisation unit', 'error'); return; }
    const downloadBtn = document.getElementById('downloadDatasetsBtn');
    if (downloadBtn) { downloadBtn.disabled = true; downloadBtn.textContent = 'Processing...'; }
    showProgress();
    clearLogs();
    setProgress(0);
    try {
        log(`📋 Processing ${selectedDataElementIds.length} selected data elements...`, 'info');
        setProgress(20);
        const dataElementMap = {};
        datasetsData.forEach(ds => {
            if (ds.dataSetElements) {
                ds.dataSetElements.forEach(dse => {
                    if (dse.dataElement && selectedDataElementIds.includes(dse.dataElement.id)) {
                        dataElementMap[dse.dataElement.id] = dse.dataElement.displayName || dse.dataElement.name || dse.dataElement.id;
                    }
                });
            }
        });
        log(`✅ Prepared ${Object.keys(dataElementMap).length} data elements for export`, 'success');
        const orgUnitMap = {};
        const orgUnitHierarchy = {};
        log('🏢 Building organisation unit hierarchy...', 'info');
        setProgress(40);
        orgUnitsData.forEach(ou => {
            orgUnitMap[ou.id] = ou.displayName;
            orgUnitHierarchy[ou.id] = { id: ou.id, name: ou.displayName, level: ou.level, parent: ou.parent ? ou.parent.id : null };
        });
        const period = getPeriodFromForm();
        if (!period) { showNotification('Please select a period', 'error'); return; }
        log('📊 Fetching data values...', 'info');
        log(`  Data Elements: ${selectedDataElementIds.length}`, 'info');
        log(`  Org Units: ${selectedOrgUnitIds.length}`, 'info');
        log(`  Period: ${period}`, 'info');
        setProgress(60);
        const dataUrl = `${state.config.instanceUrl}/api/dataValueSets.json?dataElement=${selectedDataElementIds.join(',')}&orgUnit=${selectedOrgUnitIds.join(',')}&${period}`;
        const dataResponse = await fetchWithAuth(dataUrl);
        if (!dataResponse.ok) throw new Error('Failed to fetch data values');
        const dataValues = await dataResponse.json();
        const numValues = dataValues.dataValues?.length || 0;
        log(`✅ Retrieved ${numValues} data values`, 'success');
        setProgress(80);
        if (numValues === 0) {
            showNotification('No data values found for selected criteria', 'info');
            log('ℹ️  No data found. This could mean:', 'info');
            log('   • No data entered for this period', 'info');
            log('   • Selected org units have no data', 'info');
            log('   • Data elements are not used in this period', 'info');
            return;
        }
        log('📊 Fetching category option combos...', 'info');
        const categoryOptionComboMap = {};
        const cocUrl = `${state.config.instanceUrl}/api/categoryOptionCombos.json?fields=id,displayName,name&paging=false`;
        const cocResponse = await fetchWithAuth(cocUrl);
        if (cocResponse.ok) {
            const cocData = await cocResponse.json();
            cocData.categoryOptionCombos?.forEach(coc => { categoryOptionComboMap[coc.id] = coc.displayName || coc.name || coc.id; });
            log(`  Found ${Object.keys(categoryOptionComboMap).length} category combos`, 'info');
        }
        log('📝 Converting to wide format CSV...', 'info');
        const csvData = convertToWideFormatCSV(dataValues.dataValues, dataElementMap, orgUnitMap, categoryOptionComboMap, orgUnitHierarchy);
        const blob = new Blob([csvData], { type: 'text/csv' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `dataset-export-${new Date().toISOString().split('T')[0]}.csv`;
        link.click();
        setProgress(100);
        log('\n═══════════════════════════════════════', 'success');
        log('✅ DOWNLOAD COMPLETE!', 'success');
        log('═══════════════════════════════════════', 'success');
        log(`  📊 ${numValues} data values exported`, 'success');
        log(`  📋 ${selectedDataElementIds.length} data elements`, 'success');
        log(`  🏢 ${selectedOrgUnitIds.length} organisation units`, 'success');
        showNotification('Dataset data downloaded successfully!', 'success');
    } catch (error) {
        log(`❌ ERROR: ${error.message}`, 'error');
        showNotification(`Error: ${error.message}`, 'error');
    } finally {
        if (downloadBtn) { downloadBtn.disabled = false; downloadBtn.textContent = '📥 Download Dataset Data'; }
    }
}

function convertToWideFormatCSV(dataValues, dataElementMap, orgUnitMap, categoryOptionComboMap, orgUnitHierarchy) {
    log(`Converting ${dataValues.length} data values to wide format...`, 'info');
    const groupedData = {};
    dataValues.forEach(dv => {
        const key = `${dv.orgUnit}|${dv.period}`;
        if (!groupedData[key]) {
            groupedData[key] = { orgUnit: dv.orgUnit, period: dv.period, values: {} };
        }
        const deName = dataElementMap[dv.dataElement] || dv.dataElement;
        const cocName = categoryOptionComboMap[dv.categoryOptionCombo] || dv.categoryOptionCombo || 'default';
        let fullKey = deName;
        if (cocName !== 'default' && cocName !== dv.categoryOptionCombo) {
            fullKey = `${deName} (${cocName})`;
        }
        groupedData[key].values[fullKey] = dv.value;
    });
    const allColumns = new Set();
    Object.values(groupedData).forEach(row => { Object.keys(row.values).forEach(col => allColumns.add(col)); });
    const sortedColumns = Array.from(allColumns).sort();
    let maxLevel = 0;
    Object.values(orgUnitHierarchy).forEach(ou => { if (ou.level > maxLevel) maxLevel = ou.level; });
    const hierarchyHeaders = [];
    for (let i = 1; i <= maxLevel; i++) {
        hierarchyHeaders.push(`Level ${i}`);
    }
    const headers = [...hierarchyHeaders, 'Period', ...sortedColumns];
    let csv = headers.join(',') + '\n';
    Object.values(groupedData).forEach(row => {
        const csvRow = [];
        const hierarchy = getOrgUnitHierarchy(row.orgUnit, orgUnitHierarchy);
        for (let i = 0; i < maxLevel; i++) {
            csvRow.push(escapeCSV(hierarchy[i] || ''));
        }
        csvRow.push(escapeCSV(row.period));
        sortedColumns.forEach(col => { csvRow.push(escapeCSV(row.values[col] || '')); });
        csv += csvRow.join(',') + '\n';
    });
    log('✅ Wide format CSV conversion complete', 'success');
    return csv;
}

function getOrgUnitHierarchy(orgUnitId, orgUnitHierarchy) {
    const hierarchy = [];
    let current = orgUnitHierarchy[orgUnitId];
    while (current) {
        hierarchy.unshift(current.name);
        current = current.parent ? orgUnitHierarchy[current.parent] : null;
    }
    return hierarchy;
}

function escapeCSV(value) {
    if (value === null || value === undefined) return '';
    const stringValue = String(value);
    if (stringValue.includes(',') || stringValue.includes('"') || stringValue.includes('\n')) {
        return '"' + stringValue.replace(/"/g, '""') + '"';
    }
    return stringValue;
}

document.getElementById('loginForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const instanceUrl = document.getElementById('instanceUrl').value.trim().replace(/\/$/, '');
    const username = document.getElementById('username').value.trim();
    const password = document.getElementById('password').value.trim();
    if (!instanceUrl.startsWith('http://') && !instanceUrl.startsWith('https://')) {
        showNotification('URL must start with http:// or https://', 'error');
        return;
    }
    const submitBtn = e.target.querySelector('.btn-primary');
    submitBtn.disabled = true;
    submitBtn.textContent = 'Connecting...';
    const config = { instanceUrl, username, password };
    try {
        showNotification('Connecting via Google Apps Script...', 'info');
        const testUrl = `${instanceUrl}/api/me`;
        const authHeader = 'Basic ' + btoa(`${username}:${password}`);
        const response = await makeProxyRequest(testUrl, { method: 'GET', headers: { 'Authorization': authHeader } });
        if (response.ok) {
            const userData = await response.json();
            state.config = config;
            state.isLoggedIn = true;
            document.getElementById('loginScreen').style.display = 'none';
            document.getElementById('mainContent').classList.add('show');
            const firstName = userData.firstName || '';
            const surname = userData.surname || '';
            const fullName = `${firstName} ${surname}`.trim() || 'User';
            showNotification(`✅ Connected successfully! Welcome ${fullName}`, 'success');
        } else if (response.status === 401) {
            showNotification('❌ Invalid username or password', 'error');
        } else if (response.status === 403) {
            showNotification('❌ Access forbidden', 'error');
        } else {
            showNotification(`Authentication failed (${response.status})`, 'error');
        }
    } catch (error) {
        showNotification(`Connection failed: ${error.message}`, 'error');
    } finally {
        submitBtn.disabled = false;
        submitBtn.textContent = 'Connect to DHIS2';
    }
});

document.getElementById('logoutBtn').addEventListener('click', () => {
    state.config = null;
    state.isLoggedIn = false;
    document.getElementById('mainContent').classList.remove('show');
    document.getElementById('loginScreen').style.display = 'flex';
    document.getElementById('loginForm').reset();
});

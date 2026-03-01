/* ============================================================
   Workfront Kickstart Generator
   Generates Workfront-compatible kickstart Excel spreadsheets
   ============================================================ */

// ── Object Definitions ───────────────────────────────────────
// Each entry defines a Workfront object type with its fields.
// `key`  : the Workfront API field name used in the sheet header
// `type` : text | email | number | date | select
// `options`: for select fields – array of strings OR {value, label} objects

const WORKFRONT_OBJECTS = {
  USER: {
    label: 'User',
    pluralLabel: 'Users',
    objCode: 'USER',
    sheetName: 'Import User',
    description: 'Users and their profile information',
    color: '#5C7CFA',
    icon: 'USR',
    fields: [
      { key: 'setFirstName',   label: 'First Name',    required: true,  type: 'text',   width: 120, placeholder: 'Jane' },
      { key: 'setLastName',    label: 'Last Name',     required: true,  type: 'text',   width: 120, placeholder: 'Smith' },
      { key: 'setEmailAddr',   label: 'Email Address', required: true,  type: 'email',  width: 200, placeholder: 'jane@example.com' },
      { key: 'setUsername',    label: 'Username',      required: false, type: 'text',   width: 130, placeholder: 'jsmith' },
      { key: 'setTitle',       label: 'Title',         required: false, type: 'text',   width: 160, placeholder: 'Project Manager' },
      { key: 'setIsActive',    label: 'Active',        required: false, type: 'select', width: 90,
        options: ['TRUE', 'FALSE'], default: 'TRUE' },
    ],
  },

  PROJ: {
    label: 'Project',
    pluralLabel: 'Projects',
    objCode: 'PROJ',
    sheetName: 'Import Project',
    description: 'Projects with settings and dates',
    color: '#12B886',
    icon: 'PRJ',
    fields: [
      { key: 'setName',                      label: 'Project Name',        required: true,  type: 'text',   width: 200, placeholder: 'My Project' },
      { key: 'setDescription',               label: 'Description',         required: false, type: 'text',   width: 240, placeholder: 'Project description' },
      { key: 'setStatus',                    label: 'Status',              required: false, type: 'select', width: 150,
        options: [
          { value: 'PLN', label: 'PLN – Planning' },
          { value: 'CUR', label: 'CUR – Current' },
          { value: 'CPL', label: 'CPL – Complete' },
          { value: 'DED', label: 'DED – Dead' },
          { value: 'ONH', label: 'ONH – On Hold' },
        ], default: 'PLN' },
      { key: 'setScheduledStartDate',        label: 'Planned Start',       required: false, type: 'date',   width: 130 },
      { key: 'setScheduledCompletionDate',   label: 'Planned Completion',  required: false, type: 'date',   width: 140 },
      { key: 'setPriority',                  label: 'Priority',            required: false, type: 'select', width: 120,
        options: [
          { value: '0', label: '0 – None' },
          { value: '1', label: '1 – Urgent' },
          { value: '2', label: '2 – High' },
          { value: '3', label: '3 – Normal' },
          { value: '4', label: '4 – Low' },
        ], default: '3' },
    ],
  },

  TASK: {
    label: 'Task',
    pluralLabel: 'Tasks',
    objCode: 'TASK',
    sheetName: 'Import Task',
    description: 'Tasks within projects',
    color: '#F03E3E',
    icon: 'TSK',
    fields: [
      { key: 'setName',                    label: 'Task Name',          required: true,  type: 'text',   width: 200, placeholder: 'My Task' },
      { key: 'setProjectID',               label: 'Project ID',         required: true,  type: 'text',   width: 200, placeholder: 'Workfront Project GUID' },
      { key: 'setDescription',             label: 'Description',        required: false, type: 'text',   width: 200 },
      { key: 'setStatus',                  label: 'Status',             required: false, type: 'select', width: 130,
        options: [
          { value: 'NEW', label: 'NEW – New' },
          { value: 'INP', label: 'INP – In Progress' },
          { value: 'CPL', label: 'CPL – Complete' },
        ], default: 'NEW' },
      { key: 'setDuration',                label: 'Duration',           required: false, type: 'number', width: 90,  placeholder: '1' },
      { key: 'setDurationUnit',            label: 'Duration Unit',      required: false, type: 'select', width: 120,
        options: [
          { value: 'D', label: 'D – Days' },
          { value: 'W', label: 'W – Weeks' },
          { value: 'H', label: 'H – Hours' },
          { value: 'M', label: 'M – Minutes' },
        ], default: 'D' },
      { key: 'setWorkRequired',            label: 'Planned Hours',      required: false, type: 'number', width: 110, placeholder: '8' },
      { key: 'setPlannedStartDate',        label: 'Planned Start',      required: false, type: 'date',   width: 130 },
      { key: 'setPlannedCompletionDate',   label: 'Planned Completion', required: false, type: 'date',   width: 140 },
      { key: 'setParentID',                label: 'Parent Task ID',     required: false, type: 'text',   width: 160, placeholder: 'Parent Task GUID (optional)' },
    ],
  },

  OPTASK: {
    label: 'Issue',
    pluralLabel: 'Issues',
    objCode: 'OPTASK',
    sheetName: 'Import Issue',
    description: 'Issues and requests',
    color: '#E67700',
    icon: 'ISS',
    fields: [
      { key: 'setName',                    label: 'Issue Name',         required: true,  type: 'text',   width: 200, placeholder: 'My Issue' },
      { key: 'setProjectID',               label: 'Project ID',         required: true,  type: 'text',   width: 200, placeholder: 'Workfront Project GUID' },
      { key: 'setDescription',             label: 'Description',        required: false, type: 'text',   width: 200 },
      { key: 'setStatus',                  label: 'Status',             required: false, type: 'select', width: 140,
        options: [
          { value: 'NEW', label: 'NEW – New' },
          { value: 'INP', label: 'INP – In Progress' },
          { value: 'CPL', label: 'CPL – Complete' },
          { value: 'RES', label: 'RES – Resolved' },
          { value: 'CLS', label: 'CLS – Closed' },
        ], default: 'NEW' },
      { key: 'setPriority',                label: 'Priority',           required: false, type: 'select', width: 120,
        options: [
          { value: '0', label: '0 – None' },
          { value: '1', label: '1 – Urgent' },
          { value: '2', label: '2 – High' },
          { value: '3', label: '3 – Normal' },
          { value: '4', label: '4 – Low' },
        ], default: '3' },
      { key: 'setPlannedStartDate',        label: 'Planned Start',      required: false, type: 'date',   width: 130 },
      { key: 'setPlannedCompletionDate',   label: 'Planned Completion', required: false, type: 'date',   width: 140 },
    ],
  },

  CMPY: {
    label: 'Company',
    pluralLabel: 'Companies',
    objCode: 'CMPY',
    sheetName: 'Import Company',
    description: 'Company records',
    color: '#7950F2',
    icon: 'CMP',
    fields: [
      { key: 'setName',        label: 'Company Name', required: true,  type: 'text',   width: 200, placeholder: 'Acme Corp' },
      { key: 'setDescription', label: 'Description',  required: false, type: 'text',   width: 250 },
      { key: 'setIsActive',    label: 'Active',        required: false, type: 'select', width: 90,
        options: ['TRUE', 'FALSE'], default: 'TRUE' },
    ],
  },

  GROUP: {
    label: 'Group',
    pluralLabel: 'Groups',
    objCode: 'GROUP',
    sheetName: 'Import Group',
    description: 'Groups for organizing users',
    color: '#0CA678',
    icon: 'GRP',
    fields: [
      { key: 'setName',        label: 'Group Name',  required: true,  type: 'text',   width: 200, placeholder: 'Marketing' },
      { key: 'setDescription', label: 'Description', required: false, type: 'text',   width: 250 },
      { key: 'setIsPublic',    label: 'Public',      required: false, type: 'select', width: 90,
        options: ['TRUE', 'FALSE'], default: 'FALSE' },
    ],
  },

  ROLE: {
    label: 'Job Role',
    pluralLabel: 'Job Roles',
    objCode: 'ROLE',
    sheetName: 'Import Role',
    description: 'Job roles for resource management',
    color: '#E8590C',
    icon: 'ROL',
    fields: [
      { key: 'setName',            label: 'Role Name',    required: true,  type: 'text',   width: 200, placeholder: 'Project Manager' },
      { key: 'setDescription',     label: 'Description',  required: false, type: 'text',   width: 240 },
      { key: 'setBillingPerHour',  label: 'Billing Rate', required: false, type: 'number', width: 120, placeholder: '100.00' },
      { key: 'setCostPerHour',     label: 'Cost Rate',    required: false, type: 'number', width: 120, placeholder: '75.00' },
      { key: 'setIsActive',        label: 'Active',       required: false, type: 'select', width: 90,
        options: ['TRUE', 'FALSE'], default: 'TRUE' },
    ],
  },

  PORT: {
    label: 'Portfolio',
    pluralLabel: 'Portfolios',
    objCode: 'PORT',
    sheetName: 'Import Portfolio',
    description: 'Portfolio records',
    color: '#1098AD',
    icon: 'PRT',
    fields: [
      { key: 'setName',        label: 'Portfolio Name', required: true,  type: 'text',   width: 200, placeholder: 'My Portfolio' },
      { key: 'setDescription', label: 'Description',    required: false, type: 'text',   width: 250 },
      { key: 'setStatus',      label: 'Status',         required: false, type: 'select', width: 130,
        options: [
          { value: 'IDA', label: 'IDA – Idea' },
          { value: 'CUR', label: 'CUR – Current' },
          { value: 'CPL', label: 'CPL – Complete' },
        ], default: 'IDA' },
    ],
  },

  PRGM: {
    label: 'Program',
    pluralLabel: 'Programs',
    objCode: 'PRGM',
    sheetName: 'Import Program',
    description: 'Programs within portfolios',
    color: '#D6336C',
    icon: 'PGM',
    fields: [
      { key: 'setName',        label: 'Program Name',  required: true,  type: 'text', width: 200, placeholder: 'My Program' },
      { key: 'setDescription', label: 'Description',   required: false, type: 'text', width: 250 },
      { key: 'setPortfolioID', label: 'Portfolio ID',  required: false, type: 'text', width: 200, placeholder: 'Workfront Portfolio GUID' },
    ],
  },
};

// ── App State ─────────────────────────────────────────────────
const state = {
  selectedObjects: [], // ordered list of selected objCode keys
  activeTab: null,
};

// ── Init ──────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  renderObjectGrid();
  bindNavButtons();
});

// ── Step 1: Object Selection ──────────────────────────────────
function renderObjectGrid() {
  const grid = document.getElementById('objects-grid');
  grid.innerHTML = '';

  Object.entries(WORKFRONT_OBJECTS).forEach(([key, obj]) => {
    const card = document.createElement('div');
    card.className = 'object-card';
    card.dataset.key = key;
    card.innerHTML = `
      <span class="checkmark">✓</span>
      <div class="object-icon" style="background:${obj.color}">${obj.icon}</div>
      <h3>${obj.label}</h3>
      <p>${obj.description}</p>
    `;
    card.addEventListener('click', () => toggleObject(key, card));
    grid.appendChild(card);
  });
}

function toggleObject(key, card) {
  const idx = state.selectedObjects.indexOf(key);
  if (idx === -1) {
    state.selectedObjects.push(key);
    card.classList.add('selected');
  } else {
    state.selectedObjects.splice(idx, 1);
    card.classList.remove('selected');
  }
  refreshSelectionCount();
}

function refreshSelectionCount() {
  const n = state.selectedObjects.length;
  document.getElementById('selection-count').textContent =
    n === 0 ? 'No objects selected' : `${n} object type${n > 1 ? 's' : ''} selected`;
  document.getElementById('btn-step1-next').disabled = n === 0;
}

// ── Step Navigation ───────────────────────────────────────────
function bindNavButtons() {
  document.getElementById('btn-step1-next').addEventListener('click', () => {
    buildDataEntry();
    showStep(2);
  });

  document.getElementById('btn-step2-back').addEventListener('click', () => showStep(1));

  document.getElementById('btn-step2-next').addEventListener('click', () => {
    buildSummary();
    showStep(3);
  });

  document.getElementById('btn-step3-back').addEventListener('click', () => showStep(2));

  document.getElementById('btn-start-over').addEventListener('click', () => {
    state.selectedObjects = [];
    state.activeTab = null;
    document.querySelectorAll('.object-card.selected').forEach(c => c.classList.remove('selected'));
    refreshSelectionCount();
    showStep(1);
  });

  document.getElementById('btn-download').addEventListener('click', generateAndDownload);
}

function showStep(n) {
  [1, 2, 3].forEach(i => {
    document.getElementById(`panel-step-${i}`).classList.toggle('hidden', i !== n);
    const navEl = document.getElementById(`nav-step-${i}`);
    navEl.classList.toggle('active', i === n);
    navEl.classList.toggle('completed', i < n);
  });

  // Activate step connectors
  const c12 = document.getElementById('connector-1-2');
  const c23 = document.getElementById('connector-2-3');
  if (c12) c12.classList.toggle('active', n > 1);
  if (c23) c23.classList.toggle('active', n > 2);
}

// ── Step 2: Data Entry ────────────────────────────────────────
function buildDataEntry() {
  const tabBar     = document.getElementById('tab-bar');
  const tabContent = document.getElementById('tab-content');
  tabBar.innerHTML     = '';
  tabContent.innerHTML = '';

  state.selectedObjects.forEach((key, idx) => {
    const obj = WORKFRONT_OBJECTS[key];
    const isFirst = idx === 0;

    // Tab button
    const btn = document.createElement('button');
    btn.className = `tab-btn${isFirst ? ' active' : ''}`;
    btn.dataset.tabKey = key;
    btn.id = `tab-btn-${key}`;
    btn.innerHTML = `
      <span class="tab-dot" style="background:${obj.color}"></span>
      ${obj.label}
      <span class="tab-badge" id="badge-${key}">0</span>
    `;
    btn.addEventListener('click', () => switchTab(key));
    tabBar.appendChild(btn);

    // Tab panel
    const panel = document.createElement('div');
    panel.className = `tab-panel${isFirst ? ' active' : ''}`;
    panel.id = `tab-panel-${key}`;
    panel.appendChild(buildTable(key, obj));
    tabContent.appendChild(panel);

    // Seed 3 empty rows
    for (let i = 0; i < 3; i++) addRow(key);
  });

  if (state.selectedObjects.length > 0) {
    state.activeTab = state.selectedObjects[0];
  }
}

function switchTab(key) {
  state.activeTab = key;
  document.querySelectorAll('.tab-btn').forEach(b =>
    b.classList.toggle('active', b.dataset.tabKey === key));
  document.querySelectorAll('.tab-panel').forEach(p =>
    p.classList.toggle('active', p.id === `tab-panel-${key}`));
}

function buildTable(key, obj) {
  // Toolbar
  const toolbar = document.createElement('div');
  toolbar.className = 'table-toolbar';

  const left = document.createElement('div');
  left.className = 'toolbar-left';

  const addBtn = document.createElement('button');
  addBtn.className = 'btn btn-secondary btn-sm';
  addBtn.textContent = '+ Add Row';
  addBtn.addEventListener('click', () => addRow(key));

  const hint = document.createElement('span');
  hint.className = 'table-hint';
  hint.textContent = '* Required fields';

  left.append(addBtn, hint);
  toolbar.appendChild(left);

  // Table
  const wrapper = document.createElement('div');
  wrapper.className = 'table-wrapper';

  const table = document.createElement('table');
  table.className = 'data-table';
  table.id = `table-${key}`;

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');

  const thNum = document.createElement('th');
  thNum.className = 'col-num';
  thNum.textContent = '#';
  headerRow.appendChild(thNum);

  obj.fields.forEach(f => {
    const th = document.createElement('th');
    if (f.required) th.classList.add('req');
    th.textContent = f.label;
    headerRow.appendChild(th);
  });

  const thAct = document.createElement('th');
  thAct.className = 'col-actions';
  headerRow.appendChild(thAct);

  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  tbody.id = `tbody-${key}`;
  table.appendChild(tbody);

  wrapper.appendChild(table);

  const container = document.createElement('div');
  container.append(toolbar, wrapper);
  return container;
}

function addRow(key) {
  const tbody = document.getElementById(`tbody-${key}`);
  const obj   = WORKFRONT_OBJECTS[key];
  const rowIdx = tbody.rows.length + 1;

  const tr = document.createElement('tr');

  // Row number
  const tdNum = document.createElement('td');
  tdNum.className = 'col-num';
  tdNum.textContent = rowIdx;
  tr.appendChild(tdNum);

  // Field cells
  obj.fields.forEach(field => {
    const td  = document.createElement('td');
    const inp = buildInput(field, key);
    td.appendChild(inp);
    tr.appendChild(td);
  });

  // Remove button
  const tdAct = document.createElement('td');
  tdAct.className = 'col-actions';
  const removeBtn = document.createElement('button');
  removeBtn.className = 'btn btn-danger';
  removeBtn.title = 'Remove row';
  removeBtn.textContent = '✕';
  removeBtn.addEventListener('click', () => {
    tr.remove();
    renumberRows(key);
    updateBadge(key);
  });
  tdAct.appendChild(removeBtn);
  tr.appendChild(tdAct);

  tbody.appendChild(tr);
  updateBadge(key);
}

function buildInput(field, key) {
  let el;

  if (field.type === 'select') {
    el = document.createElement('select');
    el.style.minWidth = `${field.width || 120}px`;

    // Blank placeholder for optional selects
    if (!field.required) {
      const blank = document.createElement('option');
      blank.value = '';
      blank.textContent = '— select —';
      el.appendChild(blank);
    }

    const opts = field.options || [];
    opts.forEach(opt => {
      const o = document.createElement('option');
      if (typeof opt === 'string') {
        o.value = opt;
        o.textContent = opt;
      } else {
        o.value = opt.value;
        o.textContent = opt.label;
      }
      el.appendChild(o);
    });

    if (field.default !== undefined) el.value = field.default;

  } else {
    el = document.createElement('input');
    el.type = field.type || 'text';
    el.style.minWidth = `${field.width || 120}px`;
    if (field.placeholder) el.placeholder = field.placeholder;
    if (field.default !== undefined) el.value = field.default;
  }

  el.dataset.fieldKey = field.key;

  if (field.required) {
    el.addEventListener('blur', () => {
      el.classList.toggle('invalid', !el.value.trim());
    });
    el.addEventListener('input', () => {
      if (el.value.trim()) el.classList.remove('invalid');
      updateBadge(key);
    });
  } else {
    el.addEventListener('input', () => updateBadge(key));
  }

  return el;
}

function renumberRows(key) {
  const tbody = document.getElementById(`tbody-${key}`);
  Array.from(tbody.rows).forEach((row, i) => {
    row.querySelector('.col-num').textContent = i + 1;
  });
}

function updateBadge(key) {
  const badge = document.getElementById(`badge-${key}`);
  if (badge) badge.textContent = getFilledRows(key).length;
}

// ── Step 3: Summary ───────────────────────────────────────────
function buildSummary() {
  const section = document.getElementById('summary-section');
  const keys    = state.selectedObjects;

  const totalRecords = keys.reduce((sum, k) => sum + getFilledRows(k).length, 0);
  const totalSheets  = keys.filter(k => getFilledRows(k).length > 0).length;

  let cardsHTML = '';
  keys.forEach(key => {
    const obj   = WORKFRONT_OBJECTS[key];
    const count = getFilledRows(key).length;
    const empty = count === 0
      ? '<div class="summary-empty">⚠ No data — sheet will be skipped</div>'
      : '';
    cardsHTML += `
      <div class="summary-card">
        <div class="summary-icon" style="background:${obj.color}">${obj.icon}</div>
        <div>
          <div class="summary-count">${count}</div>
          <div class="summary-label">${obj.pluralLabel}</div>
          ${empty}
        </div>
      </div>`;
  });

  section.innerHTML = `
    <h3>Export Summary</h3>
    <p class="summary-desc">
      ${totalRecords} record${totalRecords !== 1 ? 's' : ''} across
      ${totalSheets} sheet${totalSheets !== 1 ? 's' : ''}.
      Empty rows and object types with no data are excluded from the file.
    </p>
    <div class="summary-grid">${cardsHTML}</div>
  `;
}

// ── Data Collection ───────────────────────────────────────────
function getFilledRows(key) {
  const tbody = document.getElementById(`tbody-${key}`);
  if (!tbody) return [];

  const obj   = WORKFRONT_OBJECTS[key];
  const rows  = [];

  Array.from(tbody.rows).forEach(tr => {
    const rowData  = {};
    let   hasValue = false;

    tr.querySelectorAll('input, select').forEach(el => {
      const val = el.value.trim();
      rowData[el.dataset.fieldKey] = val;
      if (val) hasValue = true;
    });

    if (hasValue) rows.push(rowData);
  });

  return rows;
}

// ── Date Conversion ───────────────────────────────────────────
// HTML date inputs yield YYYY-MM-DD; Workfront expects MM/DD/YYYY.
function toWFDate(htmlDate) {
  if (!htmlDate) return '';
  const parts = htmlDate.split('-');
  if (parts.length !== 3) return htmlDate;
  const [y, m, d] = parts;
  return `${m}/${d}/${y}`;
}

// ── Excel Generation ──────────────────────────────────────────
function generateAndDownload() {
  const wb   = XLSX.utils.book_new();
  const keys = state.selectedObjects;

  const dateFieldKeys = new Set();
  Object.values(WORKFRONT_OBJECTS).forEach(obj => {
    obj.fields.filter(f => f.type === 'date').forEach(f => dateFieldKeys.add(f.key));
  });

  let sheetCount = 0;

  keys.forEach(key => {
    const obj  = WORKFRONT_OBJECTS[key];
    const rows = getFilledRows(key);
    if (rows.length === 0) return;

    // Build array-of-arrays for SheetJS
    const headerRow = [`#${obj.objCode}`, ...obj.fields.map(f => f.key)];

    const dataRows = rows.map(row => [
      obj.objCode,
      ...obj.fields.map(f => {
        const raw = row[f.key] || '';
        return dateFieldKeys.has(f.key) ? toWFDate(raw) : raw;
      }),
    ]);

    const sheetData = [headerRow, ...dataRows];
    const ws        = XLSX.utils.aoa_to_sheet(sheetData);

    // Column widths (approximate character width)
    ws['!cols'] = [
      { wch: 10 }, // objCode column
      ...obj.fields.map(f => ({ wch: Math.ceil((f.width || 130) / 7) })),
    ];

    XLSX.utils.book_append_sheet(wb, ws, obj.sheetName);
    sheetCount++;
  });

  if (sheetCount === 0) {
    alert('No data to export.\nPlease go back and add at least one record to any object type.');
    return;
  }

  const today    = new Date().toISOString().slice(0, 10);
  const filename = `workfront-kickstart-${today}.xlsx`;
  XLSX.writeFile(wb, filename);
}

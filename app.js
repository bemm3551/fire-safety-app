/**
 * Fire Safety Survey App - Improved Version with ZIP Export
 * Proteção Civil de Santo Tirso
 */

'use strict';

// ============================================================
// EQUIPMENT CONFIGURATION
// ============================================================
const EQUIPMENT_CONFIGS = {
  powder:          { id: 'powder',          label: 'Pó',                  icon: 'local_fire_department', description: 'Extintor de pó',                          category: 'extintores' },
  co2_2kg:         { id: 'co2_2kg',         label: 'CO₂ 2kg',             icon: 'local_fire_department', description: 'Extintor CO₂ 2kg',                        category: 'extintores' },
  co2_5kg:         { id: 'co2_5kg',         label: 'CO₂ 5kg',             icon: 'local_fire_department', description: 'Extintor CO₂ 5kg',                        category: 'extintores' },
  hose:            { id: 'hose',            label: 'BIATC',               icon: 'water_damage',          description: 'Bobine de Incêndio Armada Tipo Comum',    category: 'bocas' },
  hose_theater:    { id: 'hose_theater',    label: 'BIATT',               icon: 'water_damage',          description: 'Bobine de Incêndio Armada Tipo Teatro',   category: 'bocas' },
  blanket:         { id: 'blanket',         label: 'Manta',               icon: 'layers',                description: 'Manta de extinção',                       category: 'bocas' },
  cdi:             { id: 'cdi',             label: 'CDI',                 icon: 'electrical_services',   description: 'Central de Deteção de Incêndio',          category: 'detecao' },
  siren:           { id: 'siren',           label: 'Sirene',              icon: 'volume_up',             description: 'Sirene de alarme',                        category: 'detecao' },
  alarm_button:    { id: 'alarm_button',    label: 'Botão Alarme',        icon: 'notifications_active',  description: 'Botão de alarme manual',                  category: 'detecao' },
  detectors:       { id: 'detectors',       label: 'Detetores',           icon: 'sensors',               description: 'Detetores de fumo/calor',                 category: 'detecao' },
  fire_door:       { id: 'fire_door',       label: 'Portas Corta-Fogo',   icon: 'door_front',            description: 'Porta corta-fogo',                        category: 'compartimentacao' },
  permanent_block: { id: 'permanent_block', label: 'Bloco Permanente',    icon: 'build',                 description: 'Sistema fixo permanente',                 category: 'compartimentacao' },
  mobile_block:    { id: 'mobile_block',    label: 'Bloco Não Permanente',icon: 'engineering',           description: 'Sistema móvel não permanente',            category: 'compartimentacao' },
  smoke_control:   { id: 'smoke_control',   label: 'Controlo de Fumo',    icon: 'cloud',                 description: 'Sistema de controlo de fumo',             category: 'compartimentacao' },
};

const EQUIPMENT_ORDER = [
  'powder', 'co2_2kg', 'co2_5kg',
  'hose', 'hose_theater', 'blanket',
  'cdi', 'siren', 'alarm_button', 'detectors',
  'fire_door', 'permanent_block', 'mobile_block', 'smoke_control',
];

const CATEGORY_GRIDS = {
  extintores:       'extintoresGrid',
  bocas:            'bocasGrid',
  detecao:          'detecaoGrid',
  compartimentacao: 'compartimentacaoGrid',
};

// ============================================================
// STATE
// ============================================================
let state = {
  counters: Object.fromEntries(EQUIPMENT_ORDER.map(k => [k, 0])),
  photos: [],
  notes: '',
  technician: '',
  history: [],
  undoStack: [],
};

let currentViewSurveyId = null;
let editingSurveyId = null;
let confirmCallback = null;

// ============================================================
// STORAGE
// ============================================================
const STORAGE_KEY = 'fire_safety_surveys_v2';

function loadHistory() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) state.history = JSON.parse(raw);
  } catch (e) {
    state.history = [];
  }
}

function saveHistory() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.history));
  } catch (e) {
    showToast('Erro ao guardar dados localmente.', 'error');
  }
}

// ============================================================
// RENDER EQUIPMENT GRID
// ============================================================
function renderEquipmentGrids() {
  EQUIPMENT_ORDER.forEach(key => {
    const cfg = EQUIPMENT_CONFIGS[key];
    const gridId = CATEGORY_GRIDS[cfg.category];
    const grid = document.getElementById(gridId);
    if (!grid) return;

    const card = document.createElement('div');
    card.className = 'counter-card' + (state.counters[key] > 0 ? ' active' : '');
    card.id = `card-${key}`;
    card.title = cfg.description;
    card.innerHTML = `
      <span class="material-icons-round counter-card-icon">${cfg.icon}</span>
      <span class="counter-card-label">${cfg.label}</span>
      <div class="counter-controls">
        <button class="counter-btn counter-btn-minus" data-key="${key}" data-action="dec" ${state.counters[key] === 0 ? 'disabled' : ''} aria-label="Diminuir ${cfg.label}">
          <span class="material-icons-round" style="font-size:18px;pointer-events:none">remove</span>
        </button>
        <span class="counter-value" id="val-${key}">${state.counters[key]}</span>
        <button class="counter-btn counter-btn-plus" data-key="${key}" data-action="inc" aria-label="Aumentar ${cfg.label}">
          <span class="material-icons-round" style="font-size:18px;pointer-events:none">add</span>
        </button>
      </div>
    `;
    grid.appendChild(card);
  });

  // Delegate click events
  document.querySelectorAll('.counter-btn').forEach(btn => {
    btn.addEventListener('click', e => {
      const key = btn.dataset.key;
      const action = btn.dataset.action;
      if (action === 'inc') increment(key);
      else if (action === 'dec') decrement(key);
    });
  });
}

// ============================================================
// COUNTER LOGIC
// ============================================================
function increment(key) {
  state.undoStack.push({ key, delta: -1, prev: state.counters[key] });
  state.counters[key]++;
  updateCounterUI(key);
  updateStats();
  showUndoBar(`${EQUIPMENT_CONFIGS[key].label} incrementado para ${state.counters[key]}`);
}

function decrement(key) {
  if (state.counters[key] <= 0) return;
  state.undoStack.push({ key, delta: 1, prev: state.counters[key] });
  state.counters[key]--;
  updateCounterUI(key);
  updateStats();
  showUndoBar(`${EQUIPMENT_CONFIGS[key].label} decrementado para ${state.counters[key]}`);
}

function undo() {
  if (state.undoStack.length === 0) return;
  const last = state.undoStack.pop();
  state.counters[last.key] = last.prev;
  updateCounterUI(last.key);
  updateStats();
  hideUndoBar();
  showToast('Ação desfeita', 'info');
}

function updateCounterUI(key) {
  const valEl = document.getElementById(`val-${key}`);
  const card = document.getElementById(`card-${key}`);
  const decBtn = card?.querySelector('.counter-btn-minus');

  if (valEl) {
    valEl.textContent = state.counters[key];
    valEl.classList.remove('bump');
    void valEl.offsetWidth; // reflow
    valEl.classList.add('bump');
    setTimeout(() => valEl.classList.remove('bump'), 200);
  }

  if (card) {
    card.classList.toggle('active', state.counters[key] > 0);
  }

  if (decBtn) {
    decBtn.disabled = state.counters[key] === 0;
  }
}

function resetCounters() {
  showConfirm('Repor Contadores', 'Tem a certeza que deseja repor todos os contadores para zero?', () => {
    EQUIPMENT_ORDER.forEach(key => {
      state.counters[key] = 0;
      updateCounterUI(key);
    });
    state.undoStack = [];
    hideUndoBar();
    updateStats();
    showToast('Contadores repostos', 'info');
  });
}

// ============================================================
// STATS
// ============================================================
function updateStats() {
  const total = EQUIPMENT_ORDER.reduce((s, k) => s + state.counters[k], 0);
  const types = EQUIPMENT_ORDER.filter(k => state.counters[k] > 0).length;

  const totalEl = document.getElementById('statTotal');
  const typesEl = document.getElementById('statTypes');

  if (totalEl) {
    totalEl.textContent = total;
    totalEl.classList.remove('bump');
    void totalEl.offsetWidth;
    totalEl.classList.add('bump');
    setTimeout(() => totalEl.classList.remove('bump'), 200);
  }

  if (typesEl) typesEl.textContent = types;
}

function updatePhotoCount() {
  const el = document.getElementById('photoCount');
  const statEl = document.getElementById('statPhotos');
  if (el) el.textContent = state.photos.length;
  if (statEl) statEl.textContent = state.photos.length;
}

// ============================================================
// UNDO BAR
// ============================================================
let undoTimeout = null;

function showUndoBar(message) {
  const bar = document.getElementById('undoBar');
  const msg = document.getElementById('undoMessage');
  if (bar) bar.style.display = 'flex';
  if (msg) msg.textContent = message;
  clearTimeout(undoTimeout);
  undoTimeout = setTimeout(hideUndoBar, 5000);
}

function hideUndoBar() {
  const bar = document.getElementById('undoBar');
  if (bar) bar.style.display = 'none';
}

// ============================================================
// MODALS
// ============================================================
function openModal(id) {
  const overlay = document.getElementById(id);
  if (overlay) {
    overlay.classList.add('open');
    document.body.style.overflow = 'hidden';
  }
}

function closeModal(id) {
  const overlay = document.getElementById(id);
  if (overlay) {
    overlay.classList.remove('open');
    document.body.style.overflow = '';
  }
}

// Close on overlay click
document.querySelectorAll('.modal-overlay').forEach(overlay => {
  overlay.addEventListener('click', e => {
    if (e.target === overlay) closeModal(overlay.id);
  });
});

// ============================================================
// PHOTOS
// ============================================================
function renderPhotos() {
  const grid = document.getElementById('photosGrid');
  const empty = document.getElementById('emptyPhotos');
  if (!grid) return;

  grid.innerHTML = '';
  state.photos.forEach((photo, i) => {
    const item = document.createElement('div');
    item.className = 'photo-item';
    item.innerHTML = `
      <img src="${photo}" alt="Foto ${i + 1}" loading="lazy" />
      <button class="photo-remove" data-index="${i}" aria-label="Remover foto ${i + 1}">
        <span class="material-icons-round">close</span>
      </button>
    `;
    grid.appendChild(item);
  });

  grid.querySelectorAll('.photo-remove').forEach(btn => {
    btn.addEventListener('click', e => {
      e.stopPropagation();
      const idx = parseInt(btn.dataset.index);
      state.photos.splice(idx, 1);
      renderPhotos();
      updatePhotoCount();
    });
  });

  if (empty) empty.style.display = state.photos.length === 0 ? 'flex' : 'none';
  updatePhotoCount();
}

function handlePhotoInput(e) {
  const files = Array.from(e.target.files);
  const remaining = 30 - state.photos.length;
  const toProcess = files.slice(0, remaining);

  if (files.length > remaining) {
    showToast(`Máximo de 30 fotos. ${files.length - remaining} foto(s) ignorada(s).`, 'info');
  }

  let processed = 0;
  toProcess.forEach(file => {
    const reader = new FileReader();
    reader.onload = ev => {
      state.photos.push(ev.target.result);
      processed++;
      if (processed === toProcess.length) {
        renderPhotos();
        updatePhotoCount();
        showToast(`${toProcess.length} foto(s) adicionada(s)`, 'success');
      }
    };
    reader.readAsDataURL(file);
  });

  e.target.value = '';
}

// ============================================================
// NOTES
// ============================================================
function openNotesModal() {
  const ta = document.getElementById('notesTextarea');
  if (ta) ta.value = state.notes;
  updateNotesCharCount();
  openModal('notesModal');
}

function saveNotes() {
  const ta = document.getElementById('notesTextarea');
  if (ta) state.notes = ta.value.trim();
  updateNotesDot();
  closeModal('notesModal');
  showToast('Notas guardadas', 'success');
}

function updateNotesCharCount() {
  const ta = document.getElementById('notesTextarea');
  const count = document.getElementById('notesCharCount');
  if (ta && count) count.textContent = ta.value.length;
}

function updateNotesDot() {
  const dot = document.getElementById('notesDot');
  if (dot) dot.style.display = state.notes ? 'block' : 'none';
}

// ============================================================
// SAVE SURVEY
// ============================================================
function openSaveModal() {
  const locInput = document.getElementById('saveLocationInput');
  const refInput = document.getElementById('saveReferenceInput');
  const mainLoc = document.getElementById('locationInput');
  const mainRef = document.getElementById('referenceInput');
  const techInput = document.getElementById('technicianInput');
  const saveTechInput = document.getElementById('saveTechnicianInput') || createTechnicianField();

  if (locInput) locInput.value = mainLoc ? mainLoc.value : '';
  if (refInput) refInput.value = mainRef ? mainRef.value : '';
  if (saveTechInput) saveTechInput.value = techInput ? techInput.value : '';

  renderSaveSummary();
  openModal('saveModal');
}

function createTechnicianField() {
  const saveModal = document.getElementById('saveModal');
  const refInput = document.getElementById('saveReferenceInput');
  if (!saveModal || !refInput || document.getElementById('saveTechnicianInput')) return null;

  const formGroup = document.createElement('div');
  formGroup.className = 'form-group';
  formGroup.innerHTML = `
    <label class="form-label" for="saveTechnicianInput">Técnico Responsável</label>
    <input type="text" id="saveTechnicianInput" class="form-input" placeholder="Nome do técnico" maxlength="100" />
  `;
  refInput.parentElement.parentElement.insertAdjacentElement('afterend', formGroup);
  return document.getElementById('saveTechnicianInput');
}

function renderSaveSummary() {
  const summary = document.getElementById('saveSummary');
  if (!summary) return;

  const total = EQUIPMENT_ORDER.reduce((s, k) => s + state.counters[k], 0);
  const types = EQUIPMENT_ORDER.filter(k => state.counters[k] > 0).length;

  summary.innerHTML = `
    <div class="save-summary-title">Resumo do Levantamento</div>
    <div class="save-summary-grid">
      <div class="save-summary-item">
        <span class="label">Total de Equipamentos</span>
        <span class="value">${total}</span>
      </div>
      <div class="save-summary-item">
        <span class="label">Tipos Contados</span>
        <span class="value">${types}</span>
      </div>
      <div class="save-summary-item">
        <span class="label">Fotos</span>
        <span class="value">${state.photos.length}</span>
      </div>
      <div class="save-summary-item">
        <span class="label">Notas</span>
        <span class="value">${state.notes ? 'Sim' : 'Não'}</span>
      </div>
    </div>
  `;
}

function confirmSave() {
  const locInput = document.getElementById('saveLocationInput');
  const refInput = document.getElementById('saveReferenceInput');
  const techInput = document.getElementById('saveTechnicianInput');

  const location = locInput ? locInput.value.trim() : '';
  const reference = refInput ? refInput.value.trim() : '';
  const technician = techInput ? techInput.value.trim() : '';

  if (!location) {
    locInput?.classList.add('error');
    locInput?.focus();
    showToast('O nome do local é obrigatório', 'error');
    return;
  }

  locInput?.classList.remove('error');

  const now = new Date();
  const survey = {
    id: editingSurveyId || Date.now().toString(),
    location,
    reference,
    technician,
    date: now.toLocaleDateString('pt-PT'),
    time: now.toLocaleTimeString('pt-PT', { hour: '2-digit', minute: '2-digit' }),
    timestamp: now.getTime(),
    counters: { ...state.counters },
    notes: state.notes,
    photos: [...state.photos],
    totalEquipment: EQUIPMENT_ORDER.reduce((s, k) => s + state.counters[k], 0),
  };

  if (editingSurveyId) {
    const idx = state.history.findIndex(s => s.id === editingSurveyId);
    if (idx !== -1) state.history[idx] = survey;
    showToast(`Levantamento "${location}" atualizado com sucesso!`, 'success');
  } else {
    state.history.unshift(survey);
    showToast(`Levantamento "${location}" guardado com sucesso!`, 'success');
  }

  saveHistory();
  closeModal('saveModal');
  resetSurveyForm();
  renderHistory();
  updateHistoryBadge();
  editingSurveyId = null;
}

function resetSurveyForm() {
  EQUIPMENT_ORDER.forEach(key => {
    state.counters[key] = 0;
    updateCounterUI(key);
  });
  state.photos = [];
  state.notes = '';
  state.technician = '';
  state.undoStack = [];

  const locInput = document.getElementById('locationInput');
  const refInput = document.getElementById('referenceInput');
  const techInput = document.getElementById('technicianInput');
  if (locInput) locInput.value = '';
  if (refInput) refInput.value = '';
  if (techInput) techInput.value = '';

  hideUndoBar();
  updateStats();
  updatePhotoCount();
  updateNotesDot();
}

// ============================================================
// HISTORY
// ============================================================
function renderHistory(filter = '') {
  const list = document.getElementById('historyList');
  const empty = document.getElementById('emptyHistory');
  const countEl = document.getElementById('historyCount');

  if (!list) return;

  const filtered = filter
    ? state.history.filter(s =>
        s.location.toLowerCase().includes(filter.toLowerCase()) ||
        (s.reference && s.reference.toLowerCase().includes(filter.toLowerCase())) ||
        (s.technician && s.technician.toLowerCase().includes(filter.toLowerCase()))
      )
    : state.history;

  if (countEl) {
    countEl.textContent = filter
      ? `${filtered.length} de ${state.history.length} levantamentos`
      : `${state.history.length} levantamento${state.history.length !== 1 ? 's' : ''} guardado${state.history.length !== 1 ? 's' : ''}`;
  }

  list.innerHTML = '';

  if (filtered.length === 0) {
    if (empty) empty.style.display = 'flex';
    return;
  }

  if (empty) empty.style.display = 'none';

  filtered.forEach(survey => {
    const card = createHistoryCard(survey);
    list.appendChild(card);
  });
}

function createHistoryCard(survey) {
  const card = document.createElement('div');
  card.className = 'history-card';
  card.dataset.id = survey.id;

  const counted = EQUIPMENT_ORDER.filter(k => survey.counters[k] > 0);
  const chips = counted.slice(0, 4).map(k => `
    <span class="chip">
      <span class="material-icons-round">${EQUIPMENT_CONFIGS[k].icon}</span>
      ${EQUIPMENT_CONFIGS[k].label}: ${survey.counters[k]}
    </span>
  `).join('');
  const moreChips = counted.length > 4 ? `<span class="chip">+${counted.length - 4} mais</span>` : '';

  card.innerHTML = `
    <div class="history-card-header">
      <div>
        <div class="history-card-title">${escapeHtml(survey.location)}</div>
        ${survey.reference ? `<div class="history-card-ref">Ref: ${escapeHtml(survey.reference)}</div>` : ''}
        ${survey.technician ? `<div class="history-card-ref">Técnico: ${escapeHtml(survey.technician)}</div>` : ''}
      </div>
      <div class="history-card-date">
        <div>${survey.date}</div>
        <div>${survey.time}</div>
      </div>
    </div>
    <div class="history-card-stats">
      <div class="history-stat">
        <span class="material-icons-round">inventory_2</span>
        <span>Total: <strong>${survey.totalEquipment || 0}</strong></span>
      </div>
      ${survey.photos && survey.photos.length > 0 ? `
        <div class="history-stat">
          <span class="material-icons-round">photo_library</span>
          <span>Fotos: <strong>${survey.photos.length}</strong></span>
        </div>
      ` : ''}
      ${survey.notes ? `
        <div class="history-stat">
          <span class="material-icons-round">edit_note</span>
          <span>Com notas</span>
        </div>
      ` : ''}
    </div>
    ${chips || moreChips ? `<div class="history-card-chips">${chips}${moreChips}</div>` : ''}
  `;

  card.addEventListener('click', () => openViewModal(survey.id));
  return card;
}

function updateHistoryBadge() {
  const badge = document.getElementById('historyBadge');
  if (badge) {
    badge.textContent = state.history.length;
    badge.style.display = state.history.length > 0 ? 'inline-block' : 'none';
  }
}

// ============================================================
// VIEW SURVEY MODAL
// ============================================================
function openViewModal(id) {
  const survey = state.history.find(s => s.id === id);
  if (!survey) return;

  currentViewSurveyId = id;

  const title = document.getElementById('viewModalTitle');
  const body = document.getElementById('viewModalBody');

  if (title) {
    title.innerHTML = `
      <span class="material-icons-round">visibility</span>
      ${escapeHtml(survey.location)}
    `;
  }

  if (body) {
    const equipmentItems = EQUIPMENT_ORDER
      .filter(k => survey.counters[k] > 0)
      .map(k => `
        <div class="view-equipment-item">
          <span class="name">${EQUIPMENT_CONFIGS[k].label}</span>
          <span class="count">${survey.counters[k]}</span>
        </div>
      `).join('');

    const photosHtml = survey.photos && survey.photos.length > 0
      ? `<div class="view-photos-grid">
          ${survey.photos.map((p, i) => `
            <div class="view-photo">
              <img src="${p}" alt="Foto ${i + 1}" loading="lazy" />
            </div>
          `).join('')}
        </div>`
      : '<p style="color:var(--text-muted);font-size:0.85rem">Sem fotos</p>';

    body.innerHTML = `
      <div class="view-section">
        <div class="view-section-title">
          <span class="material-icons-round">info</span>
          Informações
        </div>
        <div class="view-equipment-list">
          <div class="view-equipment-item">
            <span class="name">Local</span>
            <span class="count" style="background:var(--surface-2);color:var(--text-primary)">${escapeHtml(survey.location)}</span>
          </div>
          ${survey.reference ? `
            <div class="view-equipment-item">
              <span class="name">Referência</span>
              <span class="count" style="background:var(--surface-2);color:var(--text-primary)">${escapeHtml(survey.reference)}</span>
            </div>
          ` : ''}
          ${survey.technician ? `
            <div class="view-equipment-item">
              <span class="name">Técnico</span>
              <span class="count" style="background:var(--surface-2);color:var(--text-primary)">${escapeHtml(survey.technician)}</span>
            </div>
          ` : ''}
          <div class="view-equipment-item">
            <span class="name">Data e Hora</span>
            <span class="count" style="background:var(--surface-2);color:var(--text-primary)">${survey.date} ${survey.time}</span>
          </div>
          <div class="view-equipment-item">
            <span class="name">Total de Equipamentos</span>
            <span class="count">${survey.totalEquipment || 0}</span>
          </div>
        </div>
      </div>

      ${equipmentItems ? `
        <div class="view-section">
          <div class="view-section-title">
            <span class="material-icons-round">fire_extinguisher</span>
            Equipamentos Contados
          </div>
          <div class="view-equipment-list">${equipmentItems}</div>
        </div>
      ` : ''}

      ${survey.notes ? `
        <div class="view-section">
          <div class="view-section-title">
            <span class="material-icons-round">edit_note</span>
            Notas
          </div>
          <div class="view-notes-box">${escapeHtml(survey.notes)}</div>
        </div>
      ` : ''}

      <div class="view-section">
        <div class="view-section-title">
          <span class="material-icons-round">photo_library</span>
          Fotos (${survey.photos ? survey.photos.length : 0})
        </div>
        ${photosHtml}
      </div>
    `;
  }

  openModal('viewModal');
}

function editCurrentSurvey() {
  if (!currentViewSurveyId) return;
  const survey = state.history.find(s => s.id === currentViewSurveyId);
  if (!survey) return;

  editingSurveyId = currentViewSurveyId;
  state.counters = { ...survey.counters };
  state.photos = [...survey.photos];
  state.notes = survey.notes || '';
  
  const locInput = document.getElementById('locationInput');
  const refInput = document.getElementById('referenceInput');
  const techInput = document.getElementById('technicianInput');
  
  if (locInput) locInput.value = survey.location;
  if (refInput) refInput.value = survey.reference || '';
  if (techInput) techInput.value = survey.technician || '';

  EQUIPMENT_ORDER.forEach(key => updateCounterUI(key));
  updateStats();
  updatePhotoCount();
  updateNotesDot();

  closeModal('viewModal');
  switchTab('survey');
  showToast('Levantamento carregado para edição', 'info');
}

function deleteCurrentSurvey() {
  if (!currentViewSurveyId) return;
  const survey = state.history.find(s => s.id === currentViewSurveyId);
  if (!survey) return;

  showConfirm(
    'Eliminar Levantamento',
    `Tem a certeza que deseja eliminar o levantamento de "${survey.location}"?`,
    () => {
      state.history = state.history.filter(s => s.id !== currentViewSurveyId);
      saveHistory();
      closeModal('viewModal');
      renderHistory();
      updateHistoryBadge();
      showToast('Levantamento eliminado', 'info');
      currentViewSurveyId = null;
    }
  );
}

// ============================================================
// EXPORT TO ZIP (Excel + Photos)
// ============================================================
async function exportToZip() {
  if (state.history.length === 0) {
    showToast('Não há levantamentos para exportar', 'error');
    return;
  }

  try {
    showToast('A preparar exportação...', 'info');

    // Carregar bibliotecas dinamicamente (SheetJS + JSZip)
    const scripts = [
      { src: 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.min.js', name: 'XLSX' },
      { src: 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js', name: 'JSZip' }
    ];
    
    let loadedCount = 0;
    scripts.forEach(lib => {
      const script = document.createElement('script');
      script.src = lib.src;
      script.onload = () => {
        loadedCount++;
        if (loadedCount === scripts.length) {
          performZipExport();
        }
      };
      script.onerror = () => {
        showToast(`Erro ao carregar biblioteca ${lib.name}`, 'error');
      };
      document.head.appendChild(script);
    });
  } catch (e) {
    showToast('Erro ao exportar dados', 'error');
    console.error(e);
  }
}

function performZipExport() {
  const JSZip = window.JSZip;
  const XLSX = window.XLSX;
  
  if (!JSZip || !XLSX) {
    showToast('Bibliotecas de exportação não disponíveis', 'error');
    return;
  }

  const zip = new JSZip();

  // Criar ficheiro Excel com dados usando SheetJS
  const headers = [
    'Local', 'Referência', 'Técnico', 'Data', 'Hora',
    'Pó', 'CO₂ 2kg', 'CO₂ 5kg', 'BIATC', 'BIATT', 'Manta',
    'CDI', 'Sirene', 'Botão Alarme', 'Detetores',
    'Portas Corta-Fogo', 'Bloco Permanente', 'Bloco Não Permanente', 'Controlo de Fumo',
    'Total Equipamentos', 'Nº Fotos', 'Notas'
  ];

  const rows = state.history.map(s => [
    s.location,
    s.reference || '',
    s.technician || '',
    s.date,
    s.time,
    s.counters.powder || 0,
    s.counters.co2_2kg || 0,
    s.counters.co2_5kg || 0,
    s.counters.hose || 0,
    s.counters.hose_theater || 0,
    s.counters.blanket || 0,
    s.counters.cdi || 0,
    s.counters.siren || 0,
    s.counters.alarm_button || 0,
    s.counters.detectors || 0,
    s.counters.fire_door || 0,
    s.counters.permanent_block || 0,
    s.counters.mobile_block || 0,
    s.counters.smoke_control || 0,
    s.totalEquipment || 0,
    s.photos ? s.photos.length : 0,
    s.notes || '',
  ]);

  // Criar workbook Excel com SheetJS
  const wsData = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  
  // Formatar coluna de largura
  ws['!cols'] = [
    { wch: 25 }, // Local
    { wch: 15 }, // Referência
    { wch: 20 }, // Técnico
    { wch: 12 }, // Data
    { wch: 8 },  // Hora
    { wch: 8 },  // Pó
    { wch: 10 }, // CO₂ 2kg
    { wch: 10 }, // CO₂ 5kg
    { wch: 8 },  // BIATC
    { wch: 8 },  // BIATT
    { wch: 8 },  // Manta
    { wch: 8 },  // CDI
    { wch: 8 },  // Sirene
    { wch: 12 }, // Botão Alarme
    { wch: 10 }, // Detetores
    { wch: 15 }, // Portas Corta-Fogo
    { wch: 15 }, // Bloco Permanente
    { wch: 18 }, // Bloco Não Permanente
    { wch: 15 }, // Controlo de Fumo
    { wch: 12 }, // Total
    { wch: 8 },  // Nº Fotos
    { wch: 30 }  // Notas
  ];
  
  // Formatar cabeçalho
  for (let i = 0; i < headers.length; i++) {
    const cellRef = XLSX.utils.encode_cell({ r: 0, c: i });
    ws[cellRef].s = {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '1976D2' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true }
    };
  }
  
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Levantamentos');
  
  // Converter workbook para blob
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const excelBlob = new Blob([new Uint8Array(wbout)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  
  // Adicionar ficheiro Excel ao ZIP
  zip.file('levantamentos.xlsx', excelBlob);

  // Adicionar fotos organizadas por levantamento
  let photoCount = 0;
  state.history.forEach((survey, idx) => {
    if (survey.photos && survey.photos.length > 0) {
      const folderName = `fotos/${idx + 1}_${survey.location.replace(/[\/\\:*?"<>|]/g, '_')}`;
      survey.photos.forEach((photoData, photoIdx) => {
        const base64Data = photoData.split(',')[1];
        zip.folder(folderName).file(`foto_${photoIdx + 1}.jpg`, base64Data, { base64: true });
        photoCount++;
      });
    }
  });

  // Gerar ficheiro ZIP
  zip.generateAsync({ type: 'blob' }).then(blob => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `levantamentos_${new Date().toISOString().split('T')[0]}.zip`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast(`${state.history.length} levantamento(s) com ${photoCount} foto(s) exportado(s)`, 'success');
  }).catch(err => {
    showToast('Erro ao criar ficheiro ZIP', 'error');
    console.error(err);
  });
}

// ============================================================
// TOAST
// ============================================================
let toastTimeout = null;

function showToast(message, type = 'success') {
  const toast = document.getElementById('toast');
  const msg = document.getElementById('toastMessage');
  const icon = document.getElementById('toastIcon');

  if (!toast) return;

  const icons = { success: 'check_circle', error: 'error', info: 'info' };
  if (msg) msg.textContent = message;
  if (icon) icon.textContent = icons[type] || 'info';

  toast.className = `toast ${type}`;
  toast.classList.add('show');

  clearTimeout(toastTimeout);
  toastTimeout = setTimeout(() => toast.classList.remove('show'), 3000);
}

// ============================================================
// CONFIRM DIALOG
// ============================================================
function showConfirm(title, message, onConfirm) {
  const titleEl = document.getElementById('confirmTitle');
  const msgEl = document.getElementById('confirmMessage');
  if (titleEl) titleEl.textContent = title;
  if (msgEl) msgEl.textContent = message;
  confirmCallback = onConfirm;
  openModal('confirmDialog');
}

// ============================================================
// TABS
// ============================================================
function switchTab(tab) {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tab === tab);
  });
  document.querySelectorAll('.tab-content').forEach(content => {
    content.classList.toggle('active', content.id === `tab${capitalize(tab)}Content`);
  });

  if (tab === 'history') {
    renderHistory(document.getElementById('historySearch')?.value || '');
  }
}

// ============================================================
// THEME
// ============================================================
function toggleTheme() {
  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  const newTheme = isDark ? 'light' : 'dark';
  document.documentElement.setAttribute('data-theme', newTheme);
  localStorage.setItem('fire_safety_theme', newTheme);

  const icon = document.querySelector('#themeToggle .material-icons-round');
  if (icon) icon.textContent = newTheme === 'dark' ? 'light_mode' : 'dark_mode';
}

function loadTheme() {
  const saved = localStorage.getItem('fire_safety_theme');
  const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
  const theme = saved || (prefersDark ? 'dark' : 'light');
  document.documentElement.setAttribute('data-theme', theme);

  const icon = document.querySelector('#themeToggle .material-icons-round');
  if (icon) icon.textContent = theme === 'dark' ? 'light_mode' : 'dark_mode';
}

// ============================================================
// UTILITIES
// ============================================================
function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function capitalize(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

// ============================================================
// INIT
// ============================================================
function init() {
  loadTheme();
  loadHistory();
  renderEquipmentGrids();
  updateStats();
  updatePhotoCount();
  updateHistoryBadge();

  // Tab buttons
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });

  // Theme toggle
  document.getElementById('themeToggle')?.addEventListener('click', toggleTheme);

  // Reset counters
  document.getElementById('resetCounters')?.addEventListener('click', resetCounters);

  // Undo button
  document.getElementById('undoBtn')?.addEventListener('click', undo);

  // Photos
  document.getElementById('btnPhotos')?.addEventListener('click', () => {
    renderPhotos();
    openModal('photosModal');
  });

  document.getElementById('photoInput')?.addEventListener('change', handlePhotoInput);

  document.getElementById('photosModalClose')?.addEventListener('click', () => closeModal('photosModal'));
  document.getElementById('photosModalDone')?.addEventListener('click', () => closeModal('photosModal'));

  // Notes
  document.getElementById('btnNotes')?.addEventListener('click', openNotesModal);
  document.getElementById('notesModalClose')?.addEventListener('click', () => closeModal('notesModal'));
  document.getElementById('notesModalClose2')?.addEventListener('click', () => closeModal('notesModal'));
  document.getElementById('notesSave')?.addEventListener('click', saveNotes);
  document.getElementById('notesTextarea')?.addEventListener('input', updateNotesCharCount);

  // Save survey
  document.getElementById('btnSave')?.addEventListener('click', openSaveModal);
  document.getElementById('saveModalClose')?.addEventListener('click', () => closeModal('saveModal'));
  document.getElementById('saveModalCancel')?.addEventListener('click', () => closeModal('saveModal'));
  document.getElementById('saveConfirm')?.addEventListener('click', confirmSave);

  document.getElementById('saveLocationInput')?.addEventListener('keydown', e => {
    if (e.key === 'Enter') confirmSave();
  });

  // View modal
  document.getElementById('viewModalClose')?.addEventListener('click', () => closeModal('viewModal'));
  document.getElementById('viewModalClose2')?.addEventListener('click', () => closeModal('viewModal'));
  document.getElementById('viewModalDelete')?.addEventListener('click', deleteCurrentSurvey);
  document.getElementById('viewModalEdit')?.addEventListener('click', editCurrentSurvey);

  // History search
  document.getElementById('historySearch')?.addEventListener('input', e => {
    renderHistory(e.target.value);
  });

  // Export
  document.getElementById('btnExport')?.addEventListener('click', exportToZip);

  // Confirm dialog
  document.getElementById('confirmCancel')?.addEventListener('click', () => {
    closeModal('confirmDialog');
    confirmCallback = null;
  });

  document.getElementById('confirmOk')?.addEventListener('click', () => {
    closeModal('confirmDialog');
    if (confirmCallback) {
      confirmCallback();
      confirmCallback = null;
    }
  });

  // Keyboard shortcuts
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      ['photosModal', 'notesModal', 'saveModal', 'viewModal', 'confirmDialog'].forEach(closeModal);
    }
    if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
      e.preventDefault();
      undo();
    }
  });

  // Location input sync
  document.getElementById('locationInput')?.addEventListener('input', e => {
    const saveInput = document.getElementById('saveLocationInput');
    if (saveInput && !saveInput.value) saveInput.value = e.target.value;
  });

  // Technician input sync
  document.getElementById('technicianInput')?.addEventListener('input', e => {
    state.technician = e.target.value;
  });

  console.log('🔥 Fire Safety Survey App initialized with ZIP export');
}

// Start app when DOM is ready
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}

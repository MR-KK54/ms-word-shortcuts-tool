/* -------------------------------------------------------------
 * Application Logic: WordKeys Hub
 * Functionality: Shortcut recorder, VBA Generator, ListCommands parser
 * ------------------------------------------------------------- */

// Default MS Word Commands Database
const DEFAULT_COMMANDS = [
  // File Operations
  { id: 'FileSave', name: 'FileSave', type: 'builtin', category: 'File', desc: 'Saves the active document.', shortcut: 'Ctrl+S' },
  { id: 'FileSaveAs', name: 'FileSaveAs', type: 'builtin', category: 'File', desc: 'Saves the active document under a new name.', shortcut: 'F12' },
  { id: 'FileOpen', name: 'FileOpen', type: 'builtin', category: 'File', desc: 'Opens an existing document.', shortcut: 'Ctrl+O' },
  { id: 'FileClose', name: 'FileClose', type: 'builtin', category: 'File', desc: 'Closes the active document.', shortcut: 'Ctrl+W' },
  { id: 'FilePrint', name: 'FilePrint', type: 'builtin', category: 'File', desc: 'Prints the active document.', shortcut: 'Ctrl+P' },
  { id: 'FileNew', name: 'FileNew', type: 'builtin', category: 'File', desc: 'Creates a new document.', shortcut: 'Ctrl+N' },
  
  // Editing Operations
  { id: 'EditUndo', name: 'EditUndo', type: 'builtin', category: 'Edit', desc: 'Reverses the last action.', shortcut: 'Ctrl+Z' },
  { id: 'EditRedo', name: 'EditRedo', type: 'builtin', category: 'Edit', desc: 'Redoes the last undone action.', shortcut: 'Ctrl+Y' },
  { id: 'EditCut', name: 'EditCut', type: 'builtin', category: 'Edit', desc: 'Cuts the selection to the clipboard.', shortcut: 'Ctrl+X' },
  { id: 'EditCopy', name: 'EditCopy', type: 'builtin', category: 'Edit', desc: 'Copies the selection to the clipboard.', shortcut: 'Ctrl+C' },
  { id: 'EditPaste', name: 'EditPaste', type: 'builtin', category: 'Edit', desc: 'Pastes clipboard contents.', shortcut: 'Ctrl+V' },
  { id: 'EditFind', name: 'EditFind', type: 'builtin', category: 'Edit', desc: 'Searches for text in the document.', shortcut: 'Ctrl+F' },
  { id: 'EditReplace', name: 'EditReplace', type: 'builtin', category: 'Edit', desc: 'Replaces specific text with new text.', shortcut: 'Ctrl+H' },
  { id: 'EditSelectAll', name: 'EditSelectAll', type: 'builtin', category: 'Edit', desc: 'Selects the entire document.', shortcut: 'Ctrl+A' },

  // Formatting Operations
  { id: 'Bold', name: 'Bold', type: 'builtin', category: 'Format', desc: 'Applies bold format to selected text.', shortcut: 'Ctrl+B' },
  { id: 'Italic', name: 'Italic', type: 'builtin', category: 'Format', desc: 'Applies italic format to selected text.', shortcut: 'Ctrl+I' },
  { id: 'Underline', name: 'Underline', type: 'builtin', category: 'Format', desc: 'Underlines the selected text.', shortcut: 'Ctrl+U' },
  { id: 'Subscript', name: 'Subscript', type: 'builtin', category: 'Format', desc: 'Creates subscript text.', shortcut: 'Ctrl+=' },
  { id: 'Superscript', name: 'Superscript', type: 'builtin', category: 'Format', desc: 'Creates superscript text.', shortcut: 'Ctrl+Shift+=' },
  { id: 'GrowFont', name: 'GrowFont', type: 'builtin', category: 'Format', desc: 'Increases the font size.', shortcut: 'Ctrl+]' },
  { id: 'ShrinkFont', name: 'ShrinkFont', type: 'builtin', category: 'Format', desc: 'Decreases the font size.', shortcut: 'Ctrl+[' },
  { id: 'FormatPainter', name: 'FormatPainter', type: 'builtin', category: 'Format', desc: 'Copies formatting to paste elsewhere.', shortcut: 'Ctrl+Shift+C' },
  { id: 'AlignLeft', name: 'AlignLeft', type: 'builtin', category: 'Format', desc: 'Aligns paragraphs to the left margin.', shortcut: 'Ctrl+L' },
  { id: 'AlignCenter', name: 'AlignCenter', type: 'builtin', category: 'Format', desc: 'Centers paragraphs between margins.', shortcut: 'Ctrl+E' },
  { id: 'AlignRight', name: 'AlignRight', type: 'builtin', category: 'Format', desc: 'Aligns paragraphs to the right margin.', shortcut: 'Ctrl+R' },
  { id: 'AlignJustify', name: 'AlignJustify', type: 'builtin', category: 'Format', desc: 'Justifies paragraphs between margins.', shortcut: 'Ctrl+J' },

  // Styles & Structure
  { id: 'NormalStyle', name: 'NormalStyle', type: 'builtin', category: 'Style', desc: 'Applies Normal style format.', shortcut: 'Ctrl+Shift+N' },
  { id: 'StyleHeading1', name: 'StyleHeading1', type: 'builtin', category: 'Style', desc: 'Applies Heading 1 style format.', shortcut: 'Ctrl+Alt+1' },
  { id: 'StyleHeading2', name: 'StyleHeading2', type: 'builtin', category: 'Style', desc: 'Applies Heading 2 style format.', shortcut: 'Ctrl+Alt+2' },
  { id: 'StyleHeading3', name: 'StyleHeading3', type: 'builtin', category: 'Style', desc: 'Applies Heading 3 style format.', shortcut: 'Ctrl+Alt+3' },

  // Inserting Objects
  { id: 'InsertHyperlink', name: 'InsertHyperlink', type: 'builtin', category: 'Insert', desc: 'Inserts a hyperlink.', shortcut: 'Ctrl+K' },
  { id: 'InsertPageBreak', name: 'InsertPageBreak', type: 'builtin', category: 'Insert', desc: 'Inserts a manual page break.', shortcut: 'Ctrl+Enter' },

  // View & Tools
  { id: 'ViewMacro', name: 'ViewMacro', type: 'builtin', category: 'View', desc: 'Opens the macros dialog window.', shortcut: 'Alt+F8' },
  { id: 'SpellingAndGrammar', name: 'SpellingAndGrammar', type: 'builtin', category: 'View', desc: 'Starts spelling/grammar check.', shortcut: 'F7' },
  { id: 'UpdateFields', name: 'UpdateFields', type: 'builtin', category: 'View', desc: 'Updates selected active fields.', shortcut: 'F9' },
  { id: 'ToggleFieldCodes', name: 'ToggleFieldCodes', type: 'builtin', category: 'View', desc: 'Toggles showing code vs values.', shortcut: 'Alt+F9' }
];

// Complete WdKey and Standard Keys mapping for VBA BuildKeyCode
const KEY_MAP = {
  // Letters
  'A': 'wdKeyA', 'B': 'wdKeyB', 'C': 'wdKeyC', 'D': 'wdKeyD', 'E': 'wdKeyE', 'F': 'wdKeyF', 'G': 'wdKeyG',
  'H': 'wdKeyH', 'I': 'wdKeyI', 'J': 'wdKeyJ', 'K': 'wdKeyK', 'L': 'wdKeyL', 'M': 'wdKeyM', 'N': 'wdKeyN',
  'O': 'wdKeyO', 'P': 'wdKeyP', 'Q': 'wdKeyQ', 'R': 'wdKeyR', 'S': 'wdKeyS', 'T': 'wdKeyT', 'U': 'wdKeyU',
  'V': 'wdKeyV', 'W': 'wdKeyW', 'X': 'wdKeyX', 'Y': 'wdKeyY', 'Z': 'wdKeyZ',
  
  // Digits
  '0': 'wdKey0', '1': 'wdKey1', '2': 'wdKey2', '3': 'wdKey3', '4': 'wdKey4',
  '5': 'wdKey5', '6': 'wdKey6', '7': 'wdKey7', '8': 'wdKey8', '9': 'wdKey9',
  
  // Function keys
  'F1': 'wdKeyF1', 'F2': 'wdKeyF2', 'F3': 'wdKeyF3', 'F4': 'wdKeyF4', 'F5': 'wdKeyF5', 'F6': 'wdKeyF6',
  'F7': 'wdKeyF7', 'F8': 'wdKeyF8', 'F9': 'wdKeyF9', 'F10': 'wdKeyF10', 'F11': 'wdKeyF11', 'F12': 'wdKeyF12',
  
  // Navigation & Control keys
  'BACKSPACE': 'wdKeyBackspace', 'TAB': 'wdKeyTab', 'ENTER': 'wdKeyReturn', 'ESC': 'wdKeyEsc',
  'SPACEBAR': 'wdKeySpacebar', 'SPACE': 'wdKeySpacebar', 'PAGEUP': 'wdKeyPageUp', 'PAGEDOWN': 'wdKeyPageDown',
  'END': 'wdKeyEnd', 'HOME': 'wdKeyHome', 'INSERT': 'wdKeyInsert', 'DELETE': 'wdKeyDelete',
  
  // Arrow Keys (Word maps virtual codes or vbKey constants)
  'ARROWUP': '38', 'UP': '38', 'UPARROW': '38',
  'ARROWDOWN': '40', 'DOWN': '40', 'DOWNARROW': '40',
  'ARROWLEFT': '37', 'LEFT': '37', 'LEFTARROW': '37',
  'ARROWRIGHT': '39', 'RIGHT': '39', 'RIGHTARROW': '39',

  // Punctuation & Specials
  ',': 'wdKeyComma', '<': 'wdKeyComma',
  '=': 'wdKeyEquals', '+': 'wdKeyEquals',
  '-': 'wdKeyHyphen', '_': 'wdKeyHyphen',
  '.': '190', '>': '190',
  '/': '191', '?': '191',
  ';': '186', ':': '186',
  '`': '192', '~': '192',
  '[': '219', '{': '219',
  '\\': '220', '|': '220',
  ']': '221', '}': '221',
  '\'': '222', '"': '222'
};

// Global App State
let appCommands = [];
let activeFilter = 'all';
let searchPattern = '';
let currentEditingCommand = null;
let currentRecordedKeys = { ctrl: false, alt: false, shift: false, key: '' };
let activeVbaTab = 'apply';
let parsedImportedList = [];

// Initialize App
window.addEventListener('DOMContentLoaded', () => {
  loadState();
  renderCommands();
  updateStats();
  generateVbaCode();
  
  // Setup Event Listeners
  setupNavigation();
  setupSearchAndFilters();
  setupCommandForm();
  setupShortcutRecorder();
  setupImportParser();
  setupVbaOptions();
});

// Load state from LocalStorage or pre-populate with default
function loadState() {
  const saved = localStorage.getItem('wordkeys_db');
  if (saved) {
    try {
      appCommands = JSON.parse(saved);
    } catch (e) {
      console.error('Failed to parse saved commands database. Resetting to default.', e);
      appCommands = JSON.parse(JSON.stringify(DEFAULT_COMMANDS));
    }
  } else {
    appCommands = JSON.parse(JSON.stringify(DEFAULT_COMMANDS));
  }
}

// Save state to LocalStorage
function saveState() {
  localStorage.setItem('wordkeys_db', JSON.stringify(appCommands));
  updateStats();
  generateVbaCode();
}

// Navigation between tabs
function setupNavigation() {
  const navBtns = document.querySelectorAll('.nav-item');
  const tabs = document.querySelectorAll('.tab-content');
  
  navBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      // Remove active states
      navBtns.forEach(b => b.classList.remove('active'));
      tabs.forEach(t => t.classList.remove('active'));
      
      // Set active
      btn.classList.add('active');
      const tabId = btn.getAttribute('data-tab');
      document.getElementById(tabId).classList.add('active');
    });
  });

  // Dark/Light Theme Toggle
  const themeToggle = document.getElementById('theme-toggle');
  themeToggle.addEventListener('click', () => {
    document.body.classList.toggle('light-theme');
    document.body.classList.toggle('dark-theme');
    const isLight = document.body.classList.contains('light-theme');
    themeToggle.innerHTML = isLight ? '<i class="fa-solid fa-moon"></i>' : '<i class="fa-solid fa-sun"></i>';
  });
}

// Search and Filter controls
function setupSearchAndFilters() {
  const searchInput = document.getElementById('command-search');
  searchInput.addEventListener('input', (e) => {
    searchPattern = e.target.value.toLowerCase().trim();
    renderCommands();
  });
  
  const filterPills = document.querySelectorAll('.filter-pill');
  filterPills.forEach(pill => {
    pill.addEventListener('click', () => {
      filterPills.forEach(p => p.classList.remove('active'));
      pill.classList.add('active');
      activeFilter = pill.getAttribute('data-filter');
      renderCommands();
    });
  });
}

// Statistics calculations
function updateStats() {
  document.getElementById('stats-total-commands').textContent = appCommands.length;
  
  const mapped = appCommands.filter(c => c.shortcut && c.shortcut.trim() !== '').length;
  document.getElementById('stats-custom-keys').textContent = mapped;
  
  const macros = appCommands.filter(c => c.type === 'macro').length;
  document.getElementById('stats-custom-macros').textContent = macros;
}

// Renders the main table of commands
function renderCommands() {
  const body = document.getElementById('commands-list-body');
  body.innerHTML = '';
  
  // Apply search & filters
  const filtered = appCommands.filter(cmd => {
    // Filter logic
    if (activeFilter === 'builtin' && cmd.type !== 'builtin') return false;
    if (activeFilter === 'macro' && cmd.type !== 'macro') return false;
    if (activeFilter === 'customized') {
      // Find the original default command
      const original = DEFAULT_COMMANDS.find(o => o.id === cmd.id);
      if (cmd.type === 'macro') {
        if (!cmd.shortcut) return false; // Brand new custom macro without shortcut is not customized key
      } else if (original && original.shortcut === cmd.shortcut) {
        return false; // Built-in has same shortcut as default, not customized
      }
    }
    
    // Search logic
    if (searchPattern !== '') {
      const matchName = cmd.name.toLowerCase().includes(searchPattern);
      const matchDesc = cmd.desc.toLowerCase().includes(searchPattern);
      const matchShortcut = cmd.shortcut.toLowerCase().includes(searchPattern);
      return matchName || matchDesc || matchShortcut;
    }
    
    return true;
  });
  
  if (filtered.length === 0) {
    document.getElementById('no-results').classList.remove('hidden');
    document.getElementById('commands-table').classList.add('hidden');
    return;
  }
  
  document.getElementById('no-results').classList.add('hidden');
  document.getElementById('commands-table').classList.remove('hidden');
  
  filtered.forEach(cmd => {
    const tr = document.createElement('tr');
    tr.id = `cmd-row-${cmd.id}`;
    
    // Command Name Cell
    const tdName = document.createElement('td');
    tdName.className = 'cmd-title-cell';
    const spanName = document.createElement('span');
    spanName.className = 'cmd-name-text';
    spanName.textContent = cmd.name;
    const spanCategory = document.createElement('span');
    spanCategory.className = 'badge badge-category';
    spanCategory.textContent = cmd.category;
    tdName.appendChild(spanName);
    tdName.appendChild(spanCategory);
    
    // Type Cell
    const tdType = document.createElement('td');
    const badgeType = document.createElement('span');
    badgeType.className = `badge badge-${cmd.type}`;
    badgeType.textContent = cmd.type === 'builtin' ? 'Built-in' : 'Macro';
    tdType.appendChild(badgeType);
    
    // Description Cell
    const tdDesc = document.createElement('td');
    tdDesc.className = 'cmd-desc-text';
    tdDesc.textContent = cmd.desc || 'No description provided.';
    
    // Shortcut Cell
    const tdShortcut = document.createElement('td');
    tdShortcut.appendChild(renderShortcutCaps(cmd.shortcut));
    
    // Actions Cell
    const tdActions = document.createElement('td');
    tdActions.className = 'actions-cell';
    
    const btnEdit = document.createElement('button');
    btnEdit.className = 'btn btn-secondary btn-sm';
    btnEdit.innerHTML = '<i class="fa-solid fa-pen"></i> Edit';
    btnEdit.addEventListener('click', () => openEditModal(cmd));
    
    tdActions.appendChild(btnEdit);
    
    // If it's a macro or modified builtin, allow reset or delete
    if (cmd.type === 'macro') {
      const btnDelete = document.createElement('button');
      btnDelete.className = 'btn btn-ghost btn-sm';
      btnDelete.innerHTML = '<i class="fa-solid fa-trash-can" style="color: var(--danger)"></i>';
      btnDelete.title = 'Delete Custom Macro';
      btnDelete.addEventListener('click', () => deleteCustomMacro(cmd.id));
      tdActions.appendChild(btnDelete);
    } else {
      // For built-ins, check if it's different from default
      const original = DEFAULT_COMMANDS.find(o => o.id === cmd.id);
      if (original && original.shortcut !== cmd.shortcut) {
        const btnReset = document.createElement('button');
        btnReset.className = 'btn btn-ghost btn-sm';
        btnReset.innerHTML = '<i class="fa-solid fa-arrow-rotate-left"></i> Reset';
        btnReset.title = 'Reset to Default Shortcut';
        btnReset.addEventListener('click', () => resetToDefaultShortcut(cmd.id));
        tdActions.appendChild(btnReset);
      }
    }
    
    tr.appendChild(tdName);
    tr.appendChild(tdType);
    tr.appendChild(tdDesc);
    tr.appendChild(tdShortcut);
    tr.appendChild(tdActions);
    body.appendChild(tr);
  });
}

// Utility function to render keyboard shortcut keycaps
function renderShortcutCaps(shortcutStr) {
  const container = document.createElement('div');
  container.className = 'keycaps-container';
  
  if (!shortcutStr || shortcutStr.trim() === '') {
    const span = document.createElement('span');
    span.className = 'keycap-empty';
    span.textContent = 'None';
    container.appendChild(span);
    return container;
  }
  
  const keys = shortcutStr.split('+');
  keys.forEach((key, index) => {
    const isModifier = ['ctrl', 'shift', 'alt', 'cmd'].includes(key.toLowerCase());
    
    const kbd = document.createElement('kbd');
    kbd.className = isModifier ? 'key-cap modifier' : 'key-cap';
    kbd.textContent = key;
    container.appendChild(kbd);
    
    if (index < keys.length - 1) {
      const joiner = document.createElement('span');
      joiner.className = 'key-joiner';
      joiner.textContent = '+';
      container.appendChild(joiner);
    }
  });
  
  return container;
}

// Add/Edit Custom command dialog
function setupCommandForm() {
  const modal = document.getElementById('modal-custom-command');
  const btnAdd = document.getElementById('btn-add-custom');
  const form = document.getElementById('form-custom-command');
  
  // Open modal for new command
  btnAdd.addEventListener('click', () => {
    currentEditingCommand = null;
    document.getElementById('modal-command-title').textContent = 'Add Custom Command / Macro';
    document.getElementById('edit-command-id').value = '';
    document.getElementById('cmd-name').value = '';
    document.getElementById('cmd-name').disabled = false;
    document.getElementById('cmd-type').value = 'macro';
    document.getElementById('cmd-category').value = 'Custom';
    document.getElementById('cmd-desc').value = '';
    
    updateFormShortcutDisplay('');
    modal.classList.remove('hidden');
  });
  
  // Close modals
  document.querySelectorAll('.btn-close-modal').forEach(btn => {
    btn.addEventListener('click', () => {
      modal.classList.add('hidden');
    });
  });
  
  // Submit handler
  form.addEventListener('submit', (e) => {
    e.preventDefault();
    
    const id = document.getElementById('edit-command-id').value;
    const name = document.getElementById('cmd-name').value.trim();
    const type = document.getElementById('cmd-type').value;
    const category = document.getElementById('cmd-category').value;
    const desc = document.getElementById('cmd-desc').value.trim();
    const shortcut = document.getElementById('cmd-shortcut-value').value;
    
    if (!name) return;
    
    if (currentEditingCommand) {
      // Edit mode
      const idx = appCommands.findIndex(c => c.id === id);
      if (idx !== -1) {
        appCommands[idx].category = category;
        appCommands[idx].desc = desc;
        appCommands[idx].shortcut = shortcut;
      }
    } else {
      // Create mode
      const newId = type + '_' + name.replace(/[^a-zA-Z0-9]/g, '');
      // Check duplicate
      if (appCommands.some(c => c.name.toLowerCase() === name.toLowerCase())) {
        alert('A command or macro with this name already exists.');
        return;
      }
      
      appCommands.push({
        id: newId,
        name,
        type,
        category,
        desc,
        shortcut
      });
    }
    
    saveState();
    renderCommands();
    modal.classList.add('hidden');
  });
}

function openEditModal(cmd) {
  currentEditingCommand = cmd;
  document.getElementById('modal-command-title').textContent = cmd.type === 'macro' ? 'Edit Custom Macro' : 'Edit Built-in Command';
  document.getElementById('edit-command-id').value = cmd.id;
  document.getElementById('cmd-name').value = cmd.name;
  document.getElementById('cmd-name').disabled = true; // Cannot edit original name
  document.getElementById('cmd-type').value = cmd.type;
  document.getElementById('cmd-category').value = cmd.category;
  document.getElementById('cmd-desc').value = cmd.desc;
  
  updateFormShortcutDisplay(cmd.shortcut);
  document.getElementById('modal-custom-command').classList.remove('hidden');
}

function updateFormShortcutDisplay(shortcutStr) {
  const display = document.getElementById('cmd-shortcut-display');
  const hiddenInput = document.getElementById('cmd-shortcut-value');
  
  display.innerHTML = '';
  hiddenInput.value = shortcutStr;
  
  if (shortcutStr && shortcutStr.trim() !== '') {
    display.appendChild(renderShortcutCaps(shortcutStr));
  } else {
    display.innerHTML = '<span class="placeholder-text">No shortcut assigned</span>';
  }
}

function deleteCustomMacro(id) {
  if (confirm('Are you sure you want to delete this custom macro shortcut?')) {
    appCommands = appCommands.filter(c => c.id !== id);
    saveState();
    renderCommands();
  }
}

function resetToDefaultShortcut(id) {
  const original = DEFAULT_COMMANDS.find(o => o.id === id);
  if (original) {
    const idx = appCommands.findIndex(c => c.id === id);
    if (idx !== -1) {
      appCommands[idx].shortcut = original.shortcut;
      saveState();
      renderCommands();
    }
  }
}

// Interactive Shortcut Recorder logic
function setupShortcutRecorder() {
  const modal = document.getElementById('modal-recorder');
  const btnTrigger = document.getElementById('btn-record-shortcut');
  const liveDisplay = document.getElementById('recorder-live-keys');
  const btnSave = document.getElementById('btn-save-recorded');
  const btnClear = document.getElementById('btn-clear-recorded');
  
  btnTrigger.addEventListener('click', () => {
    // Reset recorded state
    currentRecordedKeys = { ctrl: false, alt: false, shift: false, key: '' };
    updateRecorderDisplay();
    modal.classList.remove('hidden');
    
    // Add keydown listener to window
    window.addEventListener('keydown', handleRecordingKeyDown, true);
  });
  
  // Close buttons
  document.querySelectorAll('.btn-close-recorder').forEach(btn => {
    btn.addEventListener('click', () => {
      window.removeEventListener('keydown', handleRecordingKeyDown, true);
      modal.classList.add('hidden');
    });
  });
  
  btnSave.addEventListener('click', () => {
    // Assemble final string
    const parts = [];
    if (currentRecordedKeys.ctrl) parts.push('Ctrl');
    if (currentRecordedKeys.shift) parts.push('Shift');
    if (currentRecordedKeys.alt) parts.push('Alt');
    if (currentRecordedKeys.key) parts.push(currentRecordedKeys.key);
    
    const finalShortcut = parts.join('+');
    updateFormShortcutDisplay(finalShortcut);
    
    window.removeEventListener('keydown', handleRecordingKeyDown, true);
    modal.classList.add('hidden');
  });
  
  btnClear.addEventListener('click', () => {
    updateFormShortcutDisplay('');
    window.removeEventListener('keydown', handleRecordingKeyDown, true);
    modal.classList.add('hidden');
  });
}

function handleRecordingKeyDown(e) {
  e.preventDefault();
  e.stopPropagation();
  
  const key = e.key;
  
  // Toggle modifier states
  if (key === 'Control') {
    currentRecordedKeys.ctrl = true;
    updateRecorderDisplay();
    return;
  }
  if (key === 'Shift') {
    currentRecordedKeys.shift = true;
    updateRecorderDisplay();
    return;
  }
  if (key === 'Alt') {
    currentRecordedKeys.alt = true;
    updateRecorderDisplay();
    return;
  }
  if (key === 'Meta') {
    currentRecordedKeys.ctrl = true; // Map Command on MacOS to Control in Word
    updateRecorderDisplay();
    return;
  }
  
  // Map standard keys
  let keyUpper = key.toUpperCase();
  
  // Map functional descriptions
  if (key === ' ') {
    keyUpper = 'Spacebar';
  } else if (key === 'ArrowUp') {
    keyUpper = 'Up';
  } else if (key === 'ArrowDown') {
    keyUpper = 'Down';
  } else if (key === 'ArrowLeft') {
    keyUpper = 'Left';
  } else if (key === 'ArrowRight') {
    keyUpper = 'Right';
  } else if (key === 'PageUp') {
    keyUpper = 'PageUp';
  } else if (key === 'PageDown') {
    keyUpper = 'PageDown';
  } else if (key === 'Escape') {
    keyUpper = 'Esc';
  }
  
  // If it's a valid key bindable key, lock it in
  currentRecordedKeys.ctrl = e.ctrlKey || e.metaKey;
  currentRecordedKeys.shift = e.shiftKey;
  currentRecordedKeys.alt = e.altKey;
  currentRecordedKeys.key = keyUpper;
  
  updateRecorderDisplay();
}

function updateRecorderDisplay() {
  const display = document.getElementById('recorder-live-keys');
  display.innerHTML = '';
  
  // Active states in virtual keyboard
  const vkeyCtrl = document.getElementById('vkey-Ctrl');
  const vkeyAlt = document.getElementById('vkey-Alt');
  const vkeyShift = document.getElementById('vkey-Shift');
  const vkeySpace = document.getElementById('vkey-Space');
  
  vkeyCtrl.classList.toggle('pressed', currentRecordedKeys.ctrl);
  vkeyAlt.classList.toggle('pressed', currentRecordedKeys.alt);
  vkeyShift.classList.toggle('pressed', currentRecordedKeys.shift);
  vkeySpace.classList.toggle('pressed', currentRecordedKeys.key === 'Spacebar');
  
  const parts = [];
  if (currentRecordedKeys.ctrl) parts.push({ label: 'Ctrl', isMod: true });
  if (currentRecordedKeys.shift) parts.push({ label: 'Shift', isMod: true });
  if (currentRecordedKeys.alt) parts.push({ label: 'Alt', isMod: true });
  if (currentRecordedKeys.key) parts.push({ label: currentRecordedKeys.key, isMod: false });
  
  if (parts.length === 0) {
    display.innerHTML = '<span class="placeholder-text">Press a key combination...</span>';
    return;
  }
  
  parts.forEach((part, index) => {
    const kbd = document.createElement('kbd');
    kbd.className = part.isMod ? 'key-cap modifier active' : 'key-cap active';
    kbd.textContent = part.label;
    display.appendChild(kbd);
    
    if (index < parts.length - 1) {
      const joiner = document.createElement('span');
      joiner.className = 'key-joiner';
      joiner.textContent = '+';
      display.appendChild(joiner);
    }
  });
}

// VBA Export Script Generation
function generateVbaCode() {
  const useGlobalScope = document.getElementById('vba-opt-normal-template').checked;
  const clearFirst = document.getElementById('vba-opt-clear-first').checked;
  const targetCodeElement = document.getElementById('vba-code-block');
  
  const today = new Date().toISOString().split('T')[0];
  const contextName = useGlobalScope ? 'NormalTemplate' : 'ActiveDocument.AttachedTemplate';
  
  let code = `Attribute VB_Name = "WordKeysAutoInstaller"
' ====================================================================
' WordKeys Hub: MS Word Keyboard Shortcut Auto-Installer
' Generated on: ${today}
' Scope: ${useGlobalScope ? 'Global (All Word Documents)' : 'Current Template / Document Only'}
' Instructions:
'   1. Copy this script (click 'Copy VBA' button).
'   2. Open MS Word.
'   3. Press Alt + F11 to open the VBA Editor.
'   4. Click Insert -> Module.
'   5. Paste this code into the module window.
'   6. Press F5 (or click Run) to execute.
' ====================================================================

Sub ApplyWordShortcuts()
    Dim wordContext As Object
    
    ' Set target context for custom shortcuts
    CustomizationContext = ${contextName}
    
    On Error Resume Next
`;

  // Section 1: Clear existing keys if selected
  if (clearFirst) {
    code += `    
    ' --- Section 1: Clear conflicting keyboard shortcuts ---
`;
    appCommands.forEach(cmd => {
      if (cmd.shortcut && cmd.shortcut.trim() !== '') {
        const vbaKeyCode = getVbaKeyCode(cmd.shortcut);
        if (vbaKeyCode) {
          code += `    FindKey(${vbaKeyCode}).Clear\n`;
        }
      }
    });
  }

  // Section 2: Assign new keys
  code += `    
    ' --- Section 2: Install custom keybindings ---
`;
  
  appCommands.forEach(cmd => {
    if (cmd.shortcut && cmd.shortcut.trim() !== '') {
      const vbaKeyCode = getVbaKeyCode(cmd.shortcut);
      if (vbaKeyCode) {
        const keyCategory = cmd.type === 'macro' ? 'wdKeyCategoryMacro' : 'wdKeyCategoryCommand';
        code += `    KeyBindings.Add KeyCategory:=${keyCategory}, _\n`;
        code += `                    Command:="${cmd.name}", _\n`;
        code += `                    KeyCode:=${vbaKeyCode}\n\n`;
      }
    }
  });

  code += `    On Error GoTo 0
    MsgBox "Congratulations! Your MS Word keyboard shortcuts have been successfully imported/updated.", vbInformation, "WordKeys Hub Installer"
End Sub
`;

  // Handle Reset code rendering
  let resetCode = `Attribute VB_Name = "WordKeysReset"
' ====================================================================
' WordKeys Hub: Reset Customized Shortcuts
' Scope: ${useGlobalScope ? 'Global (All Word Documents)' : 'Current Template / Document Only'}
' ====================================================================

Sub ResetCustomWordShortcuts()
    ' Set target context
    CustomizationContext = ${contextName}
    
    On Error Resume Next
    
    ' Clear registered shortcut codes:
`;

  appCommands.forEach(cmd => {
    if (cmd.shortcut && cmd.shortcut.trim() !== '') {
      const vbaKeyCode = getVbaKeyCode(cmd.shortcut);
      if (vbaKeyCode) {
        resetCode += `    FindKey(${vbaKeyCode}).Clear\n`;
      }
    }
  });

  resetCode += `    
    On Error GoTo 0
    MsgBox "Customized keyboard shortcuts have been reset.", vbInformation, "WordKeys Hub Reset"
End Sub
`;

  // Render the selected tab code
  targetCodeElement.textContent = activeVbaTab === 'apply' ? code : resetCode;
}

// Convert string shortcut like "Ctrl+Shift+S" to VBA BuildKeyCode
function getVbaKeyCode(shortcutStr) {
  const keys = shortcutStr.split('+');
  const vbaArgs = [];
  
  // Sort modifiers first for cleaner code (VBA standard is Ctrl, Shift, Alt)
  const sortedKeys = [];
  const modifiers = ['ctrl', 'shift', 'alt'];
  
  // Process modifiers
  modifiers.forEach(mod => {
    const idx = keys.findIndex(k => k.toLowerCase() === mod);
    if (idx !== -1) {
      if (mod === 'ctrl') sortedKeys.push('wdKeyControl');
      if (mod === 'shift') sortedKeys.push('wdKeyShift');
      if (mod === 'alt') sortedKeys.push('wdKeyAlt');
      keys.splice(idx, 1);
    }
  });
  
  // Process primary key
  if (keys.length > 0) {
    const primaryKey = keys[0].toUpperCase();
    const mapped = KEY_MAP[primaryKey];
    if (mapped) {
      sortedKeys.push(mapped);
    } else {
      // If we don't have it in map, try to represent it directly
      sortedKeys.push(`wdKey${primaryKey}`);
    }
  }
  
  if (sortedKeys.length === 0) return null;
  return `BuildKeyCode(${sortedKeys.join(', ')})`;
}

// Hook options changes to VBA regenerator
function setupVbaOptions() {
  document.getElementById('vba-opt-normal-template').addEventListener('change', generateVbaCode);
  document.getElementById('vba-opt-clear-first').addEventListener('change', generateVbaCode);
  
  const applyTab = document.getElementById('tab-vba-apply');
  const resetTab = document.getElementById('tab-vba-reset');
  
  applyTab.addEventListener('click', () => {
    applyTab.classList.add('active');
    resetTab.classList.remove('active');
    activeVbaTab = 'apply';
    generateVbaCode();
  });
  
  resetTab.addEventListener('click', () => {
    resetTab.classList.add('active');
    applyTab.classList.remove('active');
    activeVbaTab = 'reset';
    generateVbaCode();
  });
  
  // Clipboard Copy Button
  document.getElementById('btn-copy-vba').addEventListener('click', () => {
    const code = document.getElementById('vba-code-block').textContent;
    navigator.clipboard.writeText(code).then(() => {
      const copyBtn = document.getElementById('btn-copy-vba');
      copyBtn.innerHTML = '<i class="fa-solid fa-check"></i> Copied!';
      setTimeout(() => {
        copyBtn.innerHTML = '<i class="fa-solid fa-copy"></i> Copy VBA';
      }, 2000);
    });
  });

  // Download File Button
  document.getElementById('btn-download-vba').addEventListener('click', () => {
    const code = document.getElementById('vba-code-block').textContent;
    const blob = new Blob([code], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = activeVbaTab === 'apply' ? 'WordKeys_Installer.bas' : 'WordKeys_Reset.bas';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  // JSON Backups Export
  document.getElementById('btn-backup-export').addEventListener('click', () => {
    const dataStr = JSON.stringify(appCommands, null, 2);
    const blob = new Blob([dataStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `WordKeys_Backup_${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  // JSON Backups Restore
  const backupImport = document.getElementById('backup-import-file');
  backupImport.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(evt) {
      try {
        const imported = JSON.parse(evt.target.result);
        if (Array.isArray(imported)) {
          // Rudimentary check
          const isValid = imported.every(x => x.name && x.id && x.type);
          if (isValid) {
            if (confirm(`Do you want to overwrite your active library with ${imported.length} shortcuts from the backup file?`)) {
              appCommands = imported;
              saveState();
              renderCommands();
              alert('Shortcuts database restored successfully!');
            }
          } else {
            alert('Invalid backup file format.');
          }
        } else {
          alert('Invalid backup format: Must be a list of shortcuts.');
        }
      } catch (err) {
        alert('Failed to parse backup JSON file: ' + err.message);
      }
    };
    reader.readAsText(file);
  });

  // Complete Database Reset
  document.getElementById('btn-reset-db').addEventListener('click', () => {
    if (confirm('WARNING: This will delete ALL custom commands and reset all default shortcuts back to their original settings. This cannot be undone. Are you sure?')) {
      appCommands = JSON.parse(JSON.stringify(DEFAULT_COMMANDS));
      saveState();
      renderCommands();
      alert('Database reset to defaults.');
    }
  });
}

// MS Word ListCommands parser
function setupImportParser() {
  const pasteZone = document.getElementById('paste-zone');
  const importTextarea = document.getElementById('import-textarea');
  const btnParse = document.getElementById('btn-parse-paste');
  const reviewSection = document.getElementById('parse-review-section');
  const resultsBody = document.getElementById('parsed-results-body');
  const selectAllCheckbox = document.getElementById('select-all-parsed');
  
  // Clipboard paste support directly to drag drop zone
  pasteZone.addEventListener('click', () => {
    // If not clicking inside textarea
    importTextarea.focus();
  });
  
  // Prevent focus bubbling issues
  importTextarea.addEventListener('click', (e) => {
    e.stopPropagation();
  });

  btnParse.addEventListener('click', () => {
    const val = importTextarea.value.trim();
    if (!val) {
      alert('Please paste the Word ListCommands output first.');
      return;
    }
    
    parsePastedText(val);
  });
  
  selectAllCheckbox.addEventListener('change', (e) => {
    const checked = e.target.checked;
    document.querySelectorAll('.import-row-checkbox').forEach(cb => {
      cb.checked = checked;
    });
  });
  
  // Cancel Action
  document.getElementById('btn-cancel-import').addEventListener('click', () => {
    reviewSection.classList.add('hidden');
    parsedImportedList = [];
    importTextarea.value = '';
  });
  
  // Import Action
  document.getElementById('btn-import-all').addEventListener('click', () => {
    const selectedCheckboxes = document.querySelectorAll('.import-row-checkbox:checked');
    if (selectedCheckboxes.length === 0) {
      alert('No keys selected for import.');
      return;
    }
    
    let importedCount = 0;
    selectedCheckboxes.forEach(cb => {
      const index = parseInt(cb.getAttribute('data-index'));
      const parsedItem = parsedImportedList[index];
      if (parsedItem) {
        // Find if this command already exists in active commands
        const existingIdx = appCommands.findIndex(c => c.name.toLowerCase() === parsedItem.name.toLowerCase());
        if (existingIdx !== -1) {
          // Update shortcut
          appCommands[existingIdx].shortcut = parsedItem.shortcut;
          // If category is custom and type was default, let's keep it built-in
        } else {
          // Create custom macro command in list
          appCommands.push({
            id: 'macro_' + parsedItem.name.replace(/[^a-zA-Z0-9]/g, ''),
            name: parsedItem.name,
            type: 'macro', // Since it wasn't in defaults, treat as macro
            category: 'Custom',
            desc: 'Imported from Word shortcuts table.',
            shortcut: parsedItem.shortcut
          });
        }
        importedCount++;
      }
    });
    
    saveState();
    renderCommands();
    reviewSection.classList.add('hidden');
    importTextarea.value = '';
    
    alert(`Successfully imported ${importedCount} shortcuts into your local list!`);
  });
}

function parsePastedText(text) {
  parsedImportedList = [];
  const lines = text.split(/\r?\n/);
  
  lines.forEach(line => {
    // Word ListCommands copies as Tab-separated values: CommandName [Tab] Modifiers/Keys [Tab] MenuAssignments
    // e.g. "Bold\tCtrl+B" or "Bold\tCtrl+B\tFormat"
    const cols = line.split('\t');
    if (cols.length >= 2) {
      const name = cols[0].trim();
      const rawShortcut = cols[1].trim();
      
      // Filter out header row
      if (name.toLowerCase() === 'command name' || name === '') return;
      if (rawShortcut === '' || rawShortcut.toLowerCase() === 'modifiers/shortcut keys') return;
      
      // Parse shortcut into unified format (e.g. "Ctrl + Alt + W" -> "Ctrl+Alt+W")
      const shortcutFormatted = cleanShortcutString(rawShortcut);
      if (shortcutFormatted) {
        parsedImportedList.push({
          name,
          shortcut: shortcutFormatted
        });
      }
    }
  });
  
  if (parsedImportedList.length === 0) {
    // Try space separation parsing if tab wasn't detected
    lines.forEach(line => {
      // Look for regex pattern like: WordCommandName   Ctrl+Alt+W
      const match = line.match(/^([a-zA-Z0-9_]+)\s+([CtrlAltShifCMD\d\+\-\=\[\]\\\'\/,\.;`\s]+)$/i);
      if (match) {
        const name = match[1].trim();
        const shortcut = cleanShortcutString(match[2]);
        if (name.toLowerCase() !== 'command' && shortcut) {
          parsedImportedList.push({ name, shortcut });
        }
      }
    });
  }
  
  if (parsedImportedList.length === 0) {
    alert('We could not parse any keyboard shortcuts from the pasted text. Make sure you copy the entire table directly from Word.');
    return;
  }
  
  // Render parsed list
  const resultsBody = document.getElementById('parsed-results-body');
  resultsBody.innerHTML = '';
  document.getElementById('parsed-count').textContent = parsedImportedList.length;
  
  parsedImportedList.forEach((item, index) => {
    const tr = document.createElement('tr');
    
    // Checkbox
    const tdCheck = document.createElement('td');
    tdCheck.innerHTML = `<input type="checkbox" class="import-row-checkbox" checked data-index="${index}">`;
    
    // Command Name
    const tdName = document.createElement('td');
    tdName.className = 'cmd-name-text';
    tdName.textContent = item.name;
    
    // Shortcut
    const tdShortcut = document.createElement('td');
    tdShortcut.appendChild(renderShortcutCaps(item.shortcut));
    
    // Status (Merge vs New Macro)
    const tdStatus = document.createElement('td');
    const existing = appCommands.find(c => c.name.toLowerCase() === item.name.toLowerCase());
    if (existing) {
      if (existing.shortcut === item.shortcut) {
        tdStatus.innerHTML = '<span style="color: var(--text-muted)">Identical (Already set)</span>';
      } else {
        tdStatus.innerHTML = `<span style="color: var(--warning)">Update existing (${existing.shortcut || 'None'})</span>`;
      }
    } else {
      tdStatus.innerHTML = '<span style="color: var(--success)">New custom macro</span>';
    }
    
    tr.appendChild(tdCheck);
    tr.appendChild(tdName);
    tr.appendChild(tdShortcut);
    tr.appendChild(tdStatus);
    resultsBody.appendChild(tr);
  });
  
  document.getElementById('select-all-parsed').checked = true;
  document.getElementById('parse-review-section').classList.remove('hidden');
  
  // Scroll to preview section
  document.getElementById('parse-review-section').scrollIntoView({ behavior: 'smooth' });
}

// Standardize shortcut strings from Word format (e.g. "Ctrl+Alt+Shift+Key" or "Ctrl + Alt + Key")
function cleanShortcutString(str) {
  // Word uses Alt+Ctrl+F or Alt+Ctrl+Shift+F
  // Normalize whitespace
  let clean = str.replace(/\s+/g, '');
  
  // Split keys and map modifiers properly
  const keys = clean.split('+');
  const parts = [];
  
  let hasCtrl = false;
  let hasAlt = false;
  let hasShift = false;
  let mainKey = '';
  
  keys.forEach(k => {
    const kl = k.toLowerCase();
    if (kl === 'ctrl' || kl === 'control' || kl === 'cmd') hasCtrl = true;
    else if (kl === 'alt') hasAlt = true;
    else if (kl === 'shift') hasShift = true;
    else mainKey = k;
  });
  
  if (hasCtrl) parts.push('Ctrl');
  if (hasShift) parts.push('Shift');
  if (hasAlt) parts.push('Alt');
  if (mainKey) parts.push(mainKey);
  
  if (parts.length === 0) return null;
  return parts.join('+');
}

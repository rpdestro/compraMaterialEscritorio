/** 
 * ARQUITETURA MODULAR
 */

const COLS_DEF = [
  { key: 'codigo', label: 'Código' },
  { key: 'descricao', label: 'Descrição' },
  { key: 'unidade', label: 'Unid. Medida' },
  { key: 'ataRegistro', label: 'ATA de Registro' },
  { key: 'dataAta', label: 'Data ATA', isDate: true },
  { key: 'solicitacao', label: 'Solicitação' },
  { key: 'dataSolicitacao', label: 'Data Solicitação', isDate: true },
  { key: 'pedidoCompra', label: 'Pedido de Compra' },
  { key: 'dataPedido', label: 'Data Pedido', isDate: true },
  { key: 'empenho', label: 'Empenho' },
  { key: 'solicFornecimento', label: 'Sol. Fornecimento' },
  { key: 'dataSolicFornecimento', label: 'Data Sol. Fornec.', isDate: true },
  { key: 'dataEntrega', label: 'Data Entrega', isDate: true }
];

// --- MODULE: State Management ---
const Store = {
  data: [],            
  cols: [...COLS_DEF],
  visibleCols: COLS_DEF.map(c => c.key),
  history: [],
  filters: {},       
  search: '',
  sortKey: null,
  sortAsc: true,
  page: 1,
  pageSize: 50,
  fileName: null,
  fileHandle: null,
  sheetNames: [],
  currentSheet: null,
  dirty: false,
  editingId: null,
  selectedIds: new Set(),
  activeFilterKey: null,
  wbCache: null, 
  
  setDirty(val) {
    this.dirty = val;
    UI.updateStatus();
    if (val) Storage.autoSave();
  },
  
  addHistory(action, detail) {
    this.history.unshift({ time: new Date().toLocaleString(), action, detail });
    if (this.history.length > 50) this.history.pop();
  }
};

// --- MODULE: Utilities ---
const Utils = {
  uid: () => 'r_' + Math.random().toString(36).substr(2, 9),
  $ : (sel) => document.querySelector(sel),
  $$: (sel) => Array.from(document.querySelectorAll(sel)),
  escapeHTML: (str) => String(str || '').replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m])),
  
  parseDateBR: (str) => {
    if (!str) return null;
    const m = str.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    return m ? new Date(m[3], m[2]-1, m[1]) : null;
  },
  formatDateBR: (d) => {
    if (!(d instanceof Date) || isNaN(d)) return '';
    return String(d.getDate()).padStart(2, '0') + '/' + String(d.getMonth()+1).padStart(2, '0') + '/' + d.getFullYear();
  },
  normalizeDate(val) {
    if (!val) return '';
    if (val instanceof Date) return Utils.formatDateBR(val);
    if (typeof val === 'number') {
      const p = XLSX.SSF.parse_date_code(val);
      return p ? `${String(p.d).padStart(2,'0')}/${String(p.m).padStart(2,'0')}/${p.y}` : '';
    }
    const str = String(val).trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) return Utils.formatDateBR(new Date(str + 'T00:00:00'));
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) return str;
    return str;
  },
  calcStatus(row) {
    if (row.dataEntrega) return { label: 'Entregue', class: 'status-entregue' };
    if (row.solicFornecimento || row.dataSolicFornecimento) return { label: 'Fornecedor', class: 'status-fornecedor' };
    if (row.empenho) return { label: 'Empenhado', class: 'status-empenhado' };
    if (row.pedidoCompra || row.dataPedido || row.solicitacao || row.dataSolicitacao) return { label: 'Solicitado', class: 'status-solicitado' };
    return { label: 'Pendente', class: 'status-pendente' };
  }
};

// --- MODULE: Storage & Auto-save ---
const Storage = {
  key: 'matControl_backup',
  autoSave() {
    if(!Store.data.length) return;
    const snap = { data: Store.data, fileName: Store.fileName, history: Store.history };
    localStorage.setItem(this.key, JSON.stringify(snap));
  },
  loadBackup() {
    const raw = localStorage.getItem(this.key);
    if (raw) {
      try {
        const snap = JSON.parse(raw);
        if (snap.data && snap.data.length > 0) {
          if(confirm('Foi encontrado um trabalho não salvo (Auto-save) da sua última sessão. Deseja recuperar?')) {
            Store.data = snap.data;
            Store.fileName = snap.fileName || 'backup_recuperado.xlsx';
            Store.history = snap.history || [];
            Store.setDirty(true);
            UI.initTable();
            Utils.showToast('Backup recuperado.');
            return true;
          } else {
            localStorage.removeItem(this.key);
          }
        }
      } catch(e) { console.error(e); }
    }
    return false;
  }
};

// --- MODULE: File I/O ---
const FileIO = {
  async handleFile(file, isMerge = false) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        if(!isMerge) {
          Store.wbCache = wb;
          Store.sheetNames = wb.SheetNames;
          Store.fileName = file.name;
          this.loadSheet(wb.SheetNames[0], false);
        } else {
          this.mergeSheet(wb, wb.SheetNames[0]);
        }
      } catch (err) {
        Utils.showToast('Erro ao ler a planilha.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  },
  
  loadSheet(sheetName, isTabChange = false) {
    Store.currentSheet = sheetName;
    const ws = Store.wbCache.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
    if (!raw.length) return Utils.showToast('A aba selecionada está vazia.', 'error');
    
    const headRow = raw[0];
    const missing = Store.cols.filter(c => !headRow.some(h => String(h).toLowerCase().includes(c.label.toLowerCase())));
    if (missing.length > 0 && !isTabChange) {
      alert(`Aviso: Algumas colunas padrão não foram encontradas na planilha importada:\n${missing.map(m=>m.label).join(', ')}`);
    }

    const dataRows = raw.slice(1);
    Store.data = dataRows.filter(r => r.some(c => String(c).trim() !== '')).map(r => {
      const v = { id: Utils.uid() };
      Store.cols.forEach((col, i) => {
        let val = r[i] !== undefined ? r[i] : '';
        v[col.key] = col.isDate ? Utils.normalizeDate(val) : String(val).trim();
      });
      return v;
    });

    Store.setDirty(false);
    Store.addHistory('Importação', `Carregou arquivo ${Store.fileName} (Aba: ${sheetName})`);
    UI.initTable();
  },

  mergeSheet(wb, sheetName) {
    const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: '' });
    if(raw.length < 2) return Utils.showToast('Arquivo para mesclar vazio.');
    let added = 0;
    raw.slice(1).filter(r => r.some(c => String(c).trim() !== '')).forEach(r => {
      const v = { id: Utils.uid() };
      Store.cols.forEach((col, i) => { v[col.key] = col.isDate ? Utils.normalizeDate(r[i]) : String(r[i]||'').trim(); });
      Store.data.push(v);
      added++;
    });
    Store.setDirty(true);
    Store.addHistory('Mesclagem', `Adicionou ${added} registros`);
    UI.renderBody();
    Utils.showToast(`${added} registros mesclados com sucesso!`);
  },

  buildWorkbook(onlyFiltered = false) {
    const list = onlyFiltered ? Logic.getFiltered() : Store.data;
    if(!list || list.length === 0) throw new Error('Nenhum dado para salvar.');
    
    const headerRow = [...Store.visibleCols.map(k => Store.cols.find(c=>c.key === k).label), 'Status'];
    const aoa = [headerRow].concat(list.map(r => {
      const status = Utils.calcStatus(r).label;
      return [...Store.visibleCols.map(k => r[k] || ''), status];
    }));
    
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cell_address = { c: C, r: R };
        const cell_ref = XLSX.utils.encode_cell(cell_address);
        
        if (!ws[cell_ref]) ws[cell_ref] = { t: 's', v: '' };

        const isHeader = (R === 0);
        const colName = headerRow[C];
        const isDescricao = (colName === 'Descrição');

        const cellStyle = {
          font: { bold: isHeader, color: { rgb: "000000" } },
          alignment: { 
            vertical: "center", 
            horizontal: isHeader ? "center" : (isDescricao ? "left" : "center")
          },
          border: {
            top: { style: "thin", color: { rgb: "000000" } },
            bottom: { style: "thin", color: { rgb: "000000" } },
            left: { style: "thin", color: { rgb: "000000" } },
            right: { style: "thin", color: { rgb: "000000" } }
          }
        };

        if (isHeader) {
          cellStyle.fill = { patternType: "solid", fgColor: { rgb: "DA9694" } };
        }
        ws[cell_ref].s = cellStyle;
      }
    }

    const wscols = headerRow.map(label => {
      if (label === 'Descrição') return { wch: 55 };
      if (label.includes('Código') || label.includes('Empenho') || label.includes('ATA') || label.includes('Solicitação')) return { wch: 18 };
      if (label.includes('Data')) return { wch: 14 };
      if (label === 'Status') return { wch: 12 };
      return { wch: 15 };
    });
    ws['!cols'] = wscols;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, Store.currentSheet || 'Dados');
    
    return XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  },

  async export(onlyFiltered = false) {
    try {
      const out = this.buildWorkbook(onlyFiltered);
      const blob = new Blob([out], { type: 'application/octet-stream' });
      
      let safeName = Store.fileName || 'planilha.xlsx';
      safeName = safeName.replace(/\.[a-zA-Z0-9]+$/, '') + '.xlsx';

      const isSafeToOverwrite = Store.fileHandle && Store.fileHandle.name.toLowerCase().endsWith('.xlsx');

      if (!onlyFiltered && window.showSaveFilePicker) {
        try {
          if (!isSafeToOverwrite) {
            Store.fileHandle = await window.showSaveFilePicker({
              suggestedName: safeName,
              types: [{ description: 'Planilha Excel', accept: {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']} }]
            });
            Store.fileName = Store.fileHandle.name; 
          }
          
          const writable = await Store.fileHandle.createWritable();
          await writable.write(blob);
          await writable.close();
          
          Store.setDirty(false);
          Store.addHistory('Salvamento', 'Arquivo salvo diretamente no disco com formatação');
          Utils.showToast('Arquivo atualizado com sucesso!');
          localStorage.removeItem(Storage.key);
          return;
        } catch (err) {
          if (err.name === 'AbortError') return;
          console.warn('Falha no salvamento direto. Realizando fallback para download.', err);
        }
      }

      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = (onlyFiltered ? 'filtrado_' : '') + safeName;
      a.click();
      URL.revokeObjectURL(url);
      
      if(!onlyFiltered) {
        Store.setDirty(false);
        Store.addHistory('Exportação', 'Base salva em arquivo (Download)');
        Utils.showToast('Download do arquivo concluído!');
        localStorage.removeItem(Storage.key); 
      } else {
        Utils.showToast('Visão atual exportada!');
      }
    } catch(err) {
      Utils.showToast(err.message || 'Erro ao exportar', 'error');
    }
  }
};

// --- MODULE: Core Logic (Filter & Sort) ---
const Logic = {
  getFiltered() {
    let list = Store.data;
    
    const q = Store.search.toLowerCase();
    if(q) list = list.filter(r => Store.visibleCols.some(k => String(r[k]||'').toLowerCase().includes(q)));
    
    Object.keys(Store.filters).forEach(key => {
      const filter = Store.filters[key];
      if (filter.type === 'set' && filter.values.size > 0) {
        list = list.filter(r => filter.values.has(r[key] || '(vazio)'));
      } else if (filter.type === 'date') {
        const start = filter.start ? new Date(filter.start + 'T00:00:00') : null;
        const end = filter.end ? new Date(filter.end + 'T23:59:59') : null;
        list = list.filter(r => {
          const d = Utils.parseDateBR(r[key]);
          if (!d) return false;
          if (start && d < start) return false;
          if (end && d > end) return false;
          return true;
        });
      }
    });

    if (Store.sortKey) {
      const dir = Store.sortAsc ? 1 : -1;
      list = list.slice().sort((a, b) => {
        let va = a[Store.sortKey] || '';
        let vb = b[Store.sortKey] || '';
        const isDateCol = Store.cols.find(c => c.key === Store.sortKey)?.isDate;
        
        if (isDateCol) {
          const da = Utils.parseDateBR(va), db = Utils.parseDateBR(vb);
          if (da && db) return (da.getTime() - db.getTime()) * dir;
          if (da) return -1 * dir;
          if (db) return 1 * dir;
        } else {
          const na = parseFloat(va.replace(/\./g,'').replace(',','.'));
          const nb = parseFloat(vb.replace(/\./g,'').replace(',','.'));
          if (!isNaN(na) && !isNaN(nb) && /^[\d.,]+$/.test(va)) return (na - nb) * dir;
        }
        return va.localeCompare(vb, 'pt-BR') * dir;
      });
    }
    return list;
  }
};

// --- MODULE: UI Rendering ---
const UI = {
  initTable() {
    Utils.$$('.toolbar .btn').forEach(b => b.disabled = false);
    Utils.$('#globalSearch').disabled = false;
    
    if (Store.sheetNames.length > 1) {
      const sel = Utils.$('#sheetSelect');
      sel.innerHTML = Store.sheetNames.map(s => `<option value="${s}" ${s===Store.currentSheet?'selected':''}>Aba: ${s}</option>`).join('');
      sel.style.display = 'block';
    }
    
    this.renderHead();
    this.renderBody();
  },
  
  updateStatus() {
    Utils.$('#statusDot').className = 'dot ' + (Store.dirty ? 'dirty' : 'saved');
    Utils.$('#statusText').textContent = Store.fileName ? `${Store.fileName}${Store.dirty ? ' (Não salvo)' : ''}` : 'Nenhum arquivo';
  },

  updateDashboard(filteredList) {
    const tot = filteredList.length;
    let entr = 0, emp = 0, sol = 0, pend = 0;
    
    filteredList.forEach(r => {
      const s = Utils.calcStatus(r).label;
      if (s === 'Entregue') entr++;
      else if (s === 'Empenhado' || s === 'Fornecedor') emp++;
      else if (s === 'Solicitado') sol++;
      else pend++;
    });
    
    Utils.$('#dashTotal').textContent = tot;
    Utils.$('#dashEntregues').textContent = entr;
    Utils.$('#dashEmpenhados').textContent = emp;
    Utils.$('#dashSolicitados').textContent = sol;
    Utils.$('#dashPendentes').textContent = pend;
  },

  renderHead() {
    const tr = Utils.$('#headRow');
    tr.innerHTML = '<th class="col-chk"><input type="checkbox" id="chkAll" data-testid="checkbox-all"></th>';
    
    Store.cols.filter(c => Store.visibleCols.includes(c.key)).forEach(col => {
      const th = document.createElement('th');
      const isSorted = Store.sortKey === col.key;
      const hasFilter = Store.filters[col.key] && (Store.filters[col.key].values?.size > 0 || Store.filters[col.key].start);
      if (isSorted) th.className = Store.sortAsc ? 'sort-asc' : 'sort-desc';
      
      th.innerHTML = `
        <div class="th-inner" draggable="true" data-key="${col.key}" data-testid="th-${col.key}">
          <span class="th-label">
            <svg class="sort-arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><path d="m6 15 6-6 6 6"/></svg>
            ${col.label}
          </span>
          <button class="filter-btn ${hasFilter?'active':''}" data-key="${col.key}" data-testid="btn-filter-${col.key}">
            <svg viewBox="0 0 24 24" fill="currentColor"><path d="M3 5h18l-7 8v6l-4 2v-8L3 5z"/></svg>
          </button>
        </div>`;
      tr.appendChild(th);
    });
    tr.innerHTML += '<th class="col-status">Status</th><th class="col-actions">Ações</th>';
    this.attachHeadEvents();
  },

  attachHeadEvents() {
    Utils.$$('.th-label').forEach(el => el.addEventListener('click', (e) => {
      const key = e.currentTarget.closest('.th-inner').dataset.key;
      if (Store.sortKey === key) Store.sortAsc = !Store.sortAsc;
      else { Store.sortKey = key; Store.sortAsc = true; }
      this.renderHead(); 
      this.renderBody();
    }));
    Utils.$$('.filter-btn').forEach(btn => btn.addEventListener('click', (e) => {
      e.stopPropagation();
      Modals.openFilter(btn.dataset.key, btn);
    }));
    Utils.$('#chkAll').addEventListener('change', (e) => {
      const filtered = Logic.getFiltered();
      const pageRows = filtered.slice((Store.page-1)*Store.pageSize, Store.page*Store.pageSize);
      if (e.target.checked) pageRows.forEach(r => Store.selectedIds.add(r.id));
      else pageRows.forEach(r => Store.selectedIds.delete(r.id));
      this.renderBody();
    });
    let dragKey = null;
    Utils.$$('.th-inner[draggable]').forEach(el => {
      el.addEventListener('dragstart', e => { dragKey = el.dataset.key; e.dataTransfer.effectAllowed = 'move'; });
      el.addEventListener('dragover', e => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; });
      el.addEventListener('drop', e => {
        e.preventDefault();
        const dropKey = el.dataset.key;
        if(dragKey && dragKey !== dropKey) {
          const arr = Store.visibleCols;
          arr.splice(arr.indexOf(dropKey), 0, arr.splice(arr.indexOf(dragKey), 1)[0]);
          this.renderHead(); this.renderBody();
        }
      });
    });
  },

  renderBody() {
    const filtered = Logic.getFiltered();
    this.updateDashboard(filtered);
    
    const maxPage = Math.max(1, Math.ceil(filtered.length / Store.pageSize));
    if (Store.page > maxPage) Store.page = maxPage;
    
    const start = (Store.page - 1) * Store.pageSize;
    const pageRows = filtered.slice(start, start + Store.pageSize);
    
    const tbody = Utils.$('#bodyRows');
    tbody.innerHTML = '';
    
    if(!pageRows.length) {
      tbody.innerHTML = `<tr><td colspan="100%" style="text-align:center;padding:40px;">Nenhum registro encontrado.</td></tr>`;
    } else {
      pageRows.forEach(r => {
        const isSel = Store.selectedIds.has(r.id);
        const tr = document.createElement('tr');
        if (isSel) tr.className = 'selected';
        tr.setAttribute('data-testid', `row-${r.id}`);
        
        let html = `<td class="col-chk"><input type="checkbox" class="chk-row" data-id="${r.id}" data-testid="chk-${r.id}" ${isSel?'checked':''}></td>`;
        Store.cols.filter(c => Store.visibleCols.includes(c.key)).forEach(c => {
          html += `<td>${Utils.escapeHTML(r[c.key]||'')}</td>`;
        });
        
        const status = Utils.calcStatus(r);
        html += `<td><span class="status-pill ${status.class}">${status.label}</span></td>`;
        html += `<td class="col-actions">
          <button class="row-btn clone" data-id="${r.id}" title="Duplicar" data-testid="btn-clone-${r.id}">⎘</button>
          <button class="row-btn edit" data-id="${r.id}" title="Editar" data-testid="btn-edit-${r.id}">✎</button>
          <button class="row-btn del" data-id="${r.id}" title="Excluir" data-testid="btn-del-${r.id}">✖</button>
        </td>`;
        
        tr.innerHTML = html;
        tbody.appendChild(tr);
      });
    }

    Utils.$('#rowCountText').textContent = `${filtered.length} registro(s)`;
    Utils.$('#pageIndicator').textContent = `Página ${Store.page} de ${maxPage}`;
    Utils.$('#prevPage').disabled = Store.page === 1;
    Utils.$('#nextPage').disabled = Store.page === maxPage;
    
    const bulk = Utils.$('#bulkBar');
    if (Store.selectedIds.size > 0) {
      bulk.classList.add('active');
      Utils.$('#bulkCount').textContent = `${Store.selectedIds.size} selecionados na página`;
    } else {
      bulk.classList.remove('active');
    }

    this.attachBodyEvents();
  },

  attachBodyEvents() {
    Utils.$$('.chk-row').forEach(cb => cb.addEventListener('change', e => {
      const id = e.target.dataset.id;
      if (e.target.checked) Store.selectedIds.add(id); else Store.selectedIds.delete(id);
      this.renderBody();
    }));
    Utils.$$('.row-btn.edit').forEach(b => b.addEventListener('click', e => Modals.openRecord(e.currentTarget.dataset.id)));
    Utils.$$('.row-btn.clone').forEach(b => b.addEventListener('click', e => Modals.openRecord(e.currentTarget.dataset.id, true)));
    Utils.$$('.row-btn.del').forEach(b => b.addEventListener('click', e => {
      if(confirm('Excluir este registro?')) {
        Store.data = Store.data.filter(r => r.id !== e.currentTarget.dataset.id);
        Store.setDirty(true);
        this.renderBody();
      }
    }));
  }
};

// --- MODULE: Modals & Overlays ---
const Modals = {
  init() {
    Utils.$$('.close-modal').forEach(btn => btn.addEventListener('click', (e) => {
      const overlay = e.target.closest('.overlay');
      if (overlay.id === 'recordOverlay') this.checkUnsavedRecordClose();
      else this.close(overlay);
    }));
    document.addEventListener('keydown', e => {
      if (e.key === 'Escape') {
        const openModal = Utils.$('.overlay.open');
        if (openModal) {
          if (openModal.id === 'recordOverlay') this.checkUnsavedRecordClose();
          else this.close(openModal);
        }
        Utils.$('#filterPanel').classList.remove('open');
      }
    });
  },
  close(el) { el.classList.remove('open'); },
  open(elId) { Utils.$('#'+elId).classList.add('open'); },

  initialRecordState: null,
  
  openRecord(id = null, isClone = false) {
    Store.editingId = isClone ? null : id;
    Utils.$('#recordModalTitle').textContent = id ? (isClone ? 'Clonar Registro' : 'Editar Registro') : 'Novo Registro';
    const wrap = Utils.$('#recordFormFields');
    wrap.innerHTML = '';
    
    const row = id ? Store.data.find(r => r.id === id) : {};
    
    Store.cols.forEach(col => {
      const val = row[col.key] || '';
      
      if (col.isDate) {
        wrap.innerHTML += `
          <div class="field">
            <label>${col.label}</label>
            <div style="display: flex; gap: 6px;">
              <input type="text" id="f_${col.key}" class="standard-input" value="${Utils.escapeHTML(val)}" placeholder="dd/mm/aaaa" maxlength="10" style="flex: 1;" data-testid="input-${col.key}">
              <button type="button" class="btn btn-hoje" data-target="f_${col.key}" title="Preencher com a data atual" data-testid="btn-today-${col.key}">Hoje</button>
            </div>
          </div>`;
      } else {
        wrap.innerHTML += `
          <div class="field ${col.key==='descricao'?'full':''}">
            <label>${col.label}</label>
            <input type="text" id="f_${col.key}" class="standard-input" value="${Utils.escapeHTML(val)}" data-testid="input-${col.key}">
          </div>`;
      }
    });
    
    // Binding da máscara, blur e botão "Hoje"
    Store.cols.filter(c=>c.isDate).forEach(c => {
      const inp = Utils.$('#f_'+c.key);
      
      // Autoformatação ao digitar
      inp.addEventListener('input', e => {
        let v = e.target.value.replace(/\D/g, '').slice(0, 8);
        if(v.length >= 5) {
          e.target.value = `${v.slice(0,2)}/${v.slice(2,4)}/${v.slice(4)}`;
        } else if(v.length >= 3) {
          e.target.value = `${v.slice(0,2)}/${v.slice(2)}`;
        } else {
          e.target.value = v;
        }
      });
      
      // Autocorreção do ano com dois dígitos ao sair do campo
      inp.addEventListener('blur', e => {
        let v = e.target.value;
        if (/^\d{2}\/\d{2}\/\d{2}$/.test(v)) {
          let p = v.split('/');
          // Assume prefixo '20' para anos informados com dois dígitos
          e.target.value = `${p[0]}/${p[1]}/20${p[2]}`;
        }
      });
    });

    Utils.$$('.btn-hoje').forEach(btn => {
      btn.addEventListener('click', e => {
        const target = Utils.$('#' + e.target.dataset.target);
        if (target) {
          const d = new Date();
          target.value = String(d.getDate()).padStart(2, '0') + '/' + 
                         String(d.getMonth() + 1).padStart(2, '0') + '/' + 
                         d.getFullYear();
        }
      });
    });

    this.initialRecordState = this.getRecordFormData();
    this.open('recordOverlay');
  },
  
  getRecordFormData() {
    const obj = {};
    Store.cols.forEach(c => { obj[c.key] = Utils.$('#f_'+c.key).value.trim(); });
    return obj;
  },

  checkUnsavedRecordClose() {
    const current = this.getRecordFormData();
    if (JSON.stringify(current) !== JSON.stringify(this.initialRecordState)) {
      if (!confirm('Existem alterações não salvas. Deseja realmente fechar?')) return;
    }
    this.close(Utils.$('#recordOverlay'));
  },

  saveRecord() {
    const data = this.getRecordFormData();
    if (Store.editingId) {
      const idx = Store.data.findIndex(r => r.id === Store.editingId);
      if(idx > -1) Store.data[idx] = { id: Store.editingId, ...data };
      Store.addHistory('Edição', `Atualizou registro ${data.codigo || 'sem código'}`);
    } else {
      Store.data.push({ id: Utils.uid(), ...data });
      Store.addHistory('Criação', `Novo registro ${data.codigo || 'sem código'}`);
    }
    Store.setDirty(true);
    this.initialRecordState = data; 
    this.close(Utils.$('#recordOverlay'));
    UI.renderBody();
    Utils.showToast('Salvo com sucesso!');
  },

  openFilter(key, btnRef) {
    Store.activeFilterKey = key;
    const panel = Utils.$('#filterPanel');
    const isDate = Store.cols.find(c => c.key === key)?.isDate;
    
    const datesDiv = Utils.$('#fpDates');
    if (isDate) {
      datesDiv.style.display = 'flex';
      const curFilter = Store.filters[key] || {};
      Utils.$('#fpDateStart').value = curFilter.start || '';
      Utils.$('#fpDateEnd').value = curFilter.end || '';
    } else {
      datesDiv.style.display = 'none';
    }

    const values = [...new Set(Store.data.map(r => r[key] || '(vazio)'))].sort();
    const curSet = Store.filters[key]?.values || new Set(values);
    
    const renderList = (q) => {
      const list = Utils.$('#fpList');
      list.innerHTML = '';
      values.filter(v => v.toLowerCase().includes(q.toLowerCase())).forEach(v => {
        const chk = curSet.has(v) ? 'checked' : '';
        list.innerHTML += `<label class="fp-item"><input type="checkbox" value="${Utils.escapeHTML(v)}" ${chk}><span>${Utils.escapeHTML(v)}</span></label>`;
      });
    };
    
    renderList('');
    Utils.$('#fpSearchInput').value = '';
    Utils.$('#fpSearchInput').oninput = e => renderList(e.target.value);
    
    Utils.$('#fpMarkAll').onclick = () => Utils.$$('.fp-item input').forEach(c => c.checked = true);
    Utils.$('#fpClearAll').onclick = () => Utils.$$('.fp-item input').forEach(c => c.checked = false);
    
    const rect = btnRef.getBoundingClientRect();
    panel.style.top = (rect.bottom + 5) + 'px';
    panel.style.left = Math.min(rect.left, window.innerWidth - 260) + 'px';
    panel.classList.add('open');
  },

  applyFilter() {
    const key = Store.activeFilterKey;
    const isDate = Store.cols.find(c => c.key === key)?.isDate;
    const checked = new Set(Utils.$$('.fp-item input:checked').map(c => c.value));
    const allCount = Utils.$$('.fp-item input').length;
    
    let filterObj = { type: 'set', values: checked };
    if (isDate) {
      const st = Utils.$('#fpDateStart').value;
      const nd = Utils.$('#fpDateEnd').value;
      if (st || nd) filterObj = { type: 'date', start: st, end: nd };
    }

    if (checked.size === allCount && filterObj.type === 'set') {
      delete Store.filters[key]; 
    } else {
      Store.filters[key] = filterObj;
    }
    
    Store.page = 1;
    this.close(Utils.$('#filterPanel'));
    UI.renderHead();
    UI.renderBody();
  }
};

// --- EVENTS BINDING (Power User & Toolbar) ---
document.addEventListener('DOMContentLoaded', () => {
  Storage.loadBackup();
  Modals.init();

  Utils.$('#btnTheme').addEventListener('click', () => {
    const html = document.documentElement;
    const cur = html.getAttribute('data-theme');
    html.setAttribute('data-theme', cur === 'dark' ? 'light' : 'dark');
  });

  Utils.$('#btnOpen').addEventListener('click', async () => {
    if(Store.dirty && !confirm('Existem dados não salvos. Continuar descartará as alterações. Prosseguir?')) return;
    
    if (window.showOpenFilePicker) {
      try {
        const [handle] = await window.showOpenFilePicker({
          types: [{
            description: 'Planilhas Excel e CSV',
            accept: {
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
              'application/vnd.ms-excel.sheet.macroEnabled.main+xml': ['.xlsm'],
              'application/vnd.ms-excel': ['.xls'],
              'text/csv': ['.csv']
            }
          }]
        });
        Store.fileHandle = handle;
        const file = await handle.getFile();
        FileIO.handleFile(file);
      } catch (err) {
        if (err.name !== 'AbortError') Utils.$('#fileInput').click(); // Fallback
      }
    } else {
      Utils.$('#fileInput').click(); // Fallback
    }
  });

  Utils.$('#fileInput').addEventListener('change', e => {
    if(e.target.files[0]) {
      Store.fileHandle = null; 
      FileIO.handleFile(e.target.files[0]);
    }
  });
  
  Utils.$('#btnMerge').addEventListener('click', () => Utils.$('#fileMergeInput').click());
  Utils.$('#fileMergeInput').addEventListener('change', e => {
    if(e.target.files[0]) FileIO.handleFile(e.target.files[0], true);
    e.target.value = '';
  });

  Utils.$('#btnSave').addEventListener('click', () => FileIO.export(false));
  Utils.$('#btnExportFiltered').addEventListener('click', () => FileIO.export(true));

  Utils.$('#sheetSelect').addEventListener('change', e => {
    if(Store.dirty && !confirm('Trocar de aba descartará as edições atuais. Prosseguir?')) {
      e.target.value = Store.currentSheet;
      return;
    }
    FileIO.loadSheet(e.target.value, true);
  });

  document.addEventListener('keydown', e => {
    if (e.ctrlKey && e.key.toLowerCase() === 's') { e.preventDefault(); if(!Utils.$('#btnSave').disabled) FileIO.export(false); }
    if (e.ctrlKey && e.key.toLowerCase() === 'n') { e.preventDefault(); if(!Utils.$('#btnNew').disabled) Modals.openRecord(); }
  });

  Utils.$('#btnNew').addEventListener('click', () => Modals.openRecord());
  Utils.$('#saveRecord').addEventListener('click', () => Modals.saveRecord());
  
  let searchTimeout;
  Utils.$('#globalSearch').addEventListener('input', e => {
    clearTimeout(searchTimeout);
    searchTimeout = setTimeout(() => { Store.search = e.target.value; Store.page = 1; UI.renderBody(); }, 200);
  });
  Utils.$('#btnClearFilters').addEventListener('click', () => {
    Store.filters = {}; Store.search = ''; Utils.$('#globalSearch').value = '';
    Store.page = 1; UI.renderHead(); UI.renderBody();
  });

  Utils.$('#pageSizeSelect').addEventListener('change', e => { Store.pageSize = parseInt(e.target.value); Store.page = 1; UI.renderBody(); });
  Utils.$('#prevPage').addEventListener('click', () => { if(Store.page>1){ Store.page--; UI.renderBody(); }});
  Utils.$('#nextPage').addEventListener('click', () => { Store.page++; UI.renderBody(); });

  Utils.$('#btnBulkDelete').addEventListener('click', () => {
    if(!confirm(`Excluir ${Store.selectedIds.size} registros?`)) return;
    Store.data = Store.data.filter(r => !Store.selectedIds.has(r.id));
    Store.addHistory('Exclusão Lote', `Removeu ${Store.selectedIds.size} itens`);
    Store.selectedIds.clear(); Store.setDirty(true); UI.renderBody();
  });
  Utils.$('#btnBulkEdit').addEventListener('click', () => {
    const sel = Utils.$('#bulkEditField');
    sel.innerHTML = Store.cols.map(c => `<option value="${c.key}">${c.label}</option>`).join('');
    Utils.$('#bulkEditValue').value = '';
    Modals.open('bulkEditOverlay');
  });
  Utils.$('#saveBulkEdit').addEventListener('click', () => {
    const key = Utils.$('#bulkEditField').value;
    const val = Utils.$('#bulkEditValue').value.trim();
    Store.data.forEach(r => { if(Store.selectedIds.has(r.id)) r[key] = val; });
    Store.addHistory('Edição Lote', `Alterou ${Store.cols.find(c=>c.key===key).label} em ${Store.selectedIds.size} itens`);
    Store.setDirty(true); Store.selectedIds.clear(); Modals.close(Utils.$('#bulkEditOverlay')); UI.renderBody();
  });

  Utils.$('#btnColumns').addEventListener('click', () => {
    const list = Utils.$('#colConfigList');
    list.innerHTML = Store.cols.map(c => `
      <div class="list-item">
        <span>${c.label}</span>
        <input type="checkbox" class="col-toggle" value="${c.key}" ${Store.visibleCols.includes(c.key)?'checked':''}>
      </div>
    `).join('');
    
    Utils.$$('.col-toggle').forEach(chk => chk.addEventListener('change', () => {
      Store.visibleCols = Array.from(Utils.$$('.col-toggle:checked')).map(el => el.value);
      UI.renderHead(); UI.renderBody();
    }));
    Modals.open('configOverlay');
  });

  Utils.$('#btnHistory').addEventListener('click', () => {
    const list = Utils.$('#historyList');
    list.innerHTML = Store.history.length === 0 ? '<p>Nenhuma alteração na sessão atual.</p>' :
      Store.history.map(h => `<div class="history-log"><span class="history-time">${h.time}</span><br><b>${h.action}:</b> ${h.detail}</div>`).join('');
    Modals.open('historyOverlay');
  });

  Utils.$('#fpCancel').addEventListener('click', () => Modals.close(Utils.$('#filterPanel')));
  Utils.$('#fpApply').addEventListener('click', () => Modals.applyFilter());

  window.addEventListener('beforeunload', (e) => {
    if (Store.dirty) { e.preventDefault(); e.returnValue = ''; }
  });
  
  Utils.showToast = function(msg, type='success') {
    const t = Utils.$('#toast');
    t.textContent = (type==='error'?'⚠️ ':'✅ ') + msg;
    t.className = `toast show ${type}`;
    if(t.timer) clearTimeout(t.timer);
    t.timer = setTimeout(() => t.classList.remove('show'), 3000);
  }
});
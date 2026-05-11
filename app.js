/* =========================================================
   EBD · Chamada — lógica do app
   ========================================================= */
(function () {
  'use strict';

  const STATUSES = ['P', 'A', 'J', 'F', 'FH'];
  const STATUS_LABELS = {
    P: 'Presente', A: 'Atrasado', J: 'Justificado',
    F: 'Falta', FH: 'Falta no Horário'
  };
  const STATUS_COLORS = {
    P: '#10b981', A: '#f59e0b', J: '#3b82f6', F: '#ef4444', FH: '#8b5cf6'
  };
  const MONTH_NAMES = [
    'Janeiro','Fevereiro','Março','Abril','Maio','Junho',
    'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'
  ];
  const MONTH_CACHE_KEY = 'ebd-chamada-month-v1';

  /** Estado global da aplicação. */
  const state = {
    month: '',          // 'YYYY-MM'
    date: '',           // 'YYYY-MM-DD' selecionada
    sundays: [],        // ['YYYY-MM-DD', ...]
    students: [],       // ['Nome', ...]
    grid: {},           // { 'Nome': { 'YYYY-MM-DD': 'P'|'A'|... } }
    savedDates: new Set(), // datas que já existem na planilha
    dirty: false
  };

  /** Atalhos de DOM. */
  const $ = (sel) => document.querySelector(sel);
  const el = {
    monthInput: $('#monthInput'),
    brandSubtitle: $('#brandSubtitle'),
    dateChips: $('#dateChips'),
    studentList: $('#studentList'),
    emptyState: $('#emptyState'),
    counters: { P: $('#cP'), A: $('#cA'), J: $('#cJ'), F: $('#cF'), FH: $('#cFH'), total: $('#cTotal') },
    btnSave: $('#btnSave'),
    btnHistory: $('#btnHistory'),
    btnAdd: $('#btnAdd'),
    btnExport: $('#btnExport'),
    historyModal: $('#historyModal'),
    historyTitle: $('#historyTitle'),
    historyTableWrap: $('#historyTableWrap'),
    addModal: $('#addModal'),
    addForm: $('#addForm'),
    newName: $('#newName'),
    toast: $('#toast'),
    loading: $('#loading'),
    loadingText: $('#loadingText'),
    exportSurface: $('#exportSurface')
  };

  /* ========================================================
   *  Utilitários de data
   * ======================================================== */

  function pad2(n) { return String(n).padStart(2, '0'); }

  function todayMonth() {
    const d = new Date();
    return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
  }

  function getSundaysOfMonth(month) {
    const [y, m] = month.split('-').map(Number);
    const out = [];
    const d = new Date(y, m - 1, 1);
    while (d.getMonth() === m - 1) {
      if (d.getDay() === 0) {
        out.push(`${y}-${pad2(m)}-${pad2(d.getDate())}`);
      }
      d.setDate(d.getDate() + 1);
    }
    return out;
  }

  function formatDateShort(iso) {
    const [, , dd] = iso.split('-');
    return dd;
  }

  function formatDateLong(iso) {
    const [y, m, d] = iso.split('-').map(Number);
    const names = ['domingo','segunda','terça','quarta','quinta','sexta','sábado'];
    const dt = new Date(y, m - 1, d);
    return `${pad2(d)}/${pad2(m)} — ${names[dt.getDay()]}`;
  }

  function formatMonthLong(month) {
    const [y, m] = month.split('-').map(Number);
    return `${MONTH_NAMES[m - 1]} ${y}`;
  }

  function pickClosestSunday(sundays) {
    const today = new Date();
    const todayIso = `${today.getFullYear()}-${pad2(today.getMonth()+1)}-${pad2(today.getDate())}`;
    if (sundays.includes(todayIso)) return todayIso;
    let best = sundays[0];
    let bestDiff = Infinity;
    sundays.forEach((s) => {
      const [y, m, d] = s.split('-').map(Number);
      const diff = Math.abs(new Date(y, m - 1, d) - today);
      if (diff < bestDiff) { bestDiff = diff; best = s; }
    });
    return best;
  }

  /* ========================================================
   *  API client (sem CORS preflight)
   * ======================================================== */

  function getApiUrl() {
    const url = (window.APP_CONFIG && window.APP_CONFIG.APPS_SCRIPT_URL) || '';
    if (!url || url.includes('COLE_AQUI')) {
      throw new Error('Configure a URL do Apps Script em config.js');
    }
    return url;
  }

  const API_TIMEOUT_MS = 90000;

  async function api(action, params) {
    const url = getApiUrl();
    const payload = Object.assign({ action }, params || {});
    const controller = new AbortController();
    const timeoutId = setTimeout(function () { controller.abort(); }, API_TIMEOUT_MS);
    let res;
    try {
      res = await fetch(url, {
        method: 'POST',
        // text/plain evita preflight (CORS "simple request")
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload),
        redirect: 'follow',
        signal: controller.signal
      });
    } catch (err) {
      if (err && err.name === 'AbortError') {
        throw new Error('O servidor demorou demais para responder. Tente de novo.');
      }
      throw new Error('Falha de rede. Verifique a conexão.');
    } finally {
      clearTimeout(timeoutId);
    }
    if (!res.ok) {
      throw new Error(`Servidor retornou HTTP ${res.status}`);
    }
    let json;
    try {
      json = await res.json();
    } catch (_) {
      throw new Error('Resposta inválida do servidor.');
    }
    if (!json.ok) {
      throw new Error(json.error || 'Erro desconhecido no servidor.');
    }
    return json.data;
  }

  /* ========================================================
   *  UI helpers
   * ======================================================== */

  let toastTimer = null;
  function toast(message, kind = '') {
    el.toast.textContent = message;
    el.toast.className = 'toast' + (kind ? ` toast--${kind}` : '');
    el.toast.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { el.toast.hidden = true; }, 2800);
  }

  function showLoading(text) {
    el.loadingText.textContent = text || 'Carregando…';
    el.loading.hidden = false;
  }
  function hideLoading() { el.loading.hidden = true; }

  function readMonthCache(month) {
    try {
      const raw = localStorage.getItem(MONTH_CACHE_KEY);
      if (!raw) return null;
      const obj = JSON.parse(raw);
      if (!obj || obj.month !== month || !obj.savedAt) return null;
      if (Date.now() - obj.savedAt > 7 * 24 * 60 * 60 * 1000) return null;
      return obj;
    } catch (_) {
      return null;
    }
  }

  function writeMonthCache(month, payload) {
    try {
      localStorage.setItem(
        MONTH_CACHE_KEY,
        JSON.stringify({
          month,
          savedAt: Date.now(),
          students: payload.students,
          grid: payload.grid,
          dates: payload.dates
        })
      );
    } catch (_) { /* quota / modo privado */ }
  }

  function setBrandSubtitle(text) {
    el.brandSubtitle.textContent = text;
  }

  /* ========================================================
   *  Render
   * ======================================================== */

  function renderDateChips() {
    el.dateChips.innerHTML = '';
    state.sundays.forEach((iso) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'date-chip';
      btn.dataset.date = iso;
      btn.setAttribute('aria-pressed', String(iso === state.date));
      const dd = formatDateShort(iso);
      const saved = state.savedDates.has(iso) ? ' ✓' : '';
      btn.innerHTML = `Dom ${dd}<small>${saved}</small>`;
      btn.addEventListener('click', () => {
        if (state.dirty && !confirm('Você tem alterações não salvas. Trocar de data mesmo assim?')) return;
        state.date = iso;
        state.dirty = false;
        renderDateChips();
        renderStudentList();
        renderCounters();
      });
      el.dateChips.appendChild(btn);
    });
  }

  function renderStudentList() {
    el.studentList.innerHTML = '';
    if (state.students.length === 0) {
      el.emptyState.hidden = false;
      el.studentList.setAttribute('aria-busy', 'false');
      return;
    }
    el.emptyState.hidden = true;

    const frag = document.createDocumentFragment();
    state.students.forEach((name) => {
      const current = (state.grid[name] && state.grid[name][state.date]) || '';
      const li = document.createElement('li');
      li.className = 'student';
      li.dataset.name = name;

      const header = document.createElement('h3');
      header.className = 'student__name';
      const nameSpan = document.createElement('span');
      nameSpan.textContent = name;
      const badge = document.createElement('span');
      badge.className = 'student__badge';
      badge.dataset.status = current;
      badge.textContent = current || '—';
      header.appendChild(nameSpan);
      header.appendChild(badge);

      const group = document.createElement('div');
      group.className = 'status-group';
      group.setAttribute('role', 'radiogroup');
      group.setAttribute('aria-label', `Status de ${name}`);

      STATUSES.forEach((s) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'status-btn' + (current === s ? ' is-active' : '');
        btn.dataset.status = s;
        btn.setAttribute('aria-pressed', String(current === s));
        btn.setAttribute('aria-label', `${STATUS_LABELS[s]} para ${name}`);
        btn.textContent = s;
        btn.addEventListener('click', () => setStatus(name, s));
        group.appendChild(btn);
      });

      li.appendChild(header);
      li.appendChild(group);
      frag.appendChild(li);
    });

    el.studentList.appendChild(frag);
    el.studentList.setAttribute('aria-busy', 'false');
  }

  function setStatus(name, status) {
    if (!state.date) {
      toast('Selecione uma data primeiro.', 'error');
      return;
    }
    if (!state.grid[name]) state.grid[name] = {};
    const prev = state.grid[name][state.date] || '';
    const next = prev === status ? '' : status; // tocar de novo limpa
    state.grid[name][state.date] = next;
    state.dirty = true;

    // Atualiza apenas a linha do aluno
    const row = el.studentList.querySelector(`li[data-name="${cssEscape(name)}"]`);
    if (row) {
      row.querySelectorAll('.status-btn').forEach((btn) => {
        const on = btn.dataset.status === next;
        btn.classList.toggle('is-active', on);
        btn.setAttribute('aria-pressed', String(on));
      });
      const badge = row.querySelector('.student__badge');
      if (badge) {
        badge.dataset.status = next;
        badge.textContent = next || '—';
      }
    }
    renderCounters();
  }

  function renderCounters() {
    const counts = { P: 0, A: 0, J: 0, F: 0, FH: 0, total: 0 };
    state.students.forEach((name) => {
      const s = (state.grid[name] && state.grid[name][state.date]) || '';
      if (counts[s] !== undefined) counts[s]++;
    });
    counts.total = state.students.length;
    STATUSES.forEach((s) => { el.counters[s].textContent = counts[s]; });
    el.counters.total.textContent = counts.total;

    if (state.date) {
      setBrandSubtitle(`${formatMonthLong(state.month)} · ${formatDateLong(state.date)}`);
    } else {
      setBrandSubtitle(formatMonthLong(state.month));
    }
  }

  function cssEscape(value) {
    if (window.CSS && CSS.escape) return CSS.escape(value);
    return String(value).replace(/["\\]/g, '\\$&');
  }

  /* ========================================================
   *  Histórico
   * ======================================================== */

  function renderHistory() {
    el.historyTitle.textContent = `Histórico — ${formatMonthLong(state.month)}`;

    const dates = state.sundays.slice(); // mostra todos os domingos do mês
    const wrap = el.historyTableWrap;
    wrap.innerHTML = '';

    if (state.students.length === 0) {
      wrap.innerHTML = '<p style="color:#64748b">Sem alunos cadastrados.</p>';
      return;
    }

    const table = document.createElement('table');
    table.className = 'history-table';

    const thead = document.createElement('thead');
    const trh = document.createElement('tr');
    const thName = document.createElement('th');
    thName.textContent = 'Aluno';
    trh.appendChild(thName);
    dates.forEach((d) => {
      const th = document.createElement('th');
      th.innerHTML = `Dom <b>${formatDateShort(d)}</b>`;
      trh.appendChild(th);
    });
    const thTotal = document.createElement('th');
    thTotal.textContent = '% Pres.';
    trh.appendChild(thTotal);
    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    state.students.forEach((name) => {
      const tr = document.createElement('tr');
      const tdName = document.createElement('td');
      tdName.textContent = name;
      tdName.title = name;
      tr.appendChild(tdName);

      let presentLike = 0, considered = 0;
      dates.forEach((d) => {
        const s = (state.grid[name] && state.grid[name][d]) || '';
        const td = document.createElement('td');
        if (s) {
          const span = document.createElement('span');
          span.className = 'history-cell';
          span.dataset.s = s;
          span.textContent = s;
          td.appendChild(span);
          considered++;
          if (s === 'P' || s === 'A') presentLike++;
        } else {
          td.innerHTML = '<span class="history-cell" data-s="">—</span>';
        }
        tr.appendChild(td);
      });

      const tdPct = document.createElement('td');
      tdPct.textContent = considered
        ? `${Math.round((presentLike / considered) * 100)}%`
        : '—';
      tr.appendChild(tdPct);

      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrap.appendChild(table);
  }

  /* ========================================================
   *  Ações
   * ======================================================== */

  async function loadMonth(month, opts) {
    opts = opts || {};
    state.month = month;
    state.sundays = getSundaysOfMonth(month);
    el.monthInput.value = month;
    setBrandSubtitle(formatMonthLong(month));

    const cached = !opts.skipCache ? readMonthCache(month) : null;
    let showedCache = false;
    if (cached && Array.isArray(cached.students)) {
      state.students = cached.students;
      state.grid = cached.grid && typeof cached.grid === 'object' ? cached.grid : {};
      state.savedDates = new Set(Array.isArray(cached.dates) ? cached.dates : []);
      state.students.forEach((n) => { if (!state.grid[n]) state.grid[n] = {}; });
      const preferred = opts.preserveDate && state.sundays.includes(state.date)
        ? state.date
        : pickClosestSunday(state.sundays);
      state.date = preferred || '';
      state.dirty = false;
      renderDateChips();
      renderStudentList();
      renderCounters();
      hideLoading();
      showedCache = true;
      setBrandSubtitle(`${formatMonthLong(month)} · sincronizando…`);
    } else {
      showLoading('Carregando chamada…');
    }

    try {
      const data = await api('getMonth', { month });
      state.students = data.students || [];
      state.grid = data.grid || {};
      state.savedDates = new Set(data.dates || []);
      state.students.forEach((n) => { if (!state.grid[n]) state.grid[n] = {}; });

      const preferred = opts.preserveDate && state.sundays.includes(state.date)
        ? state.date
        : pickClosestSunday(state.sundays);
      state.date = preferred || '';
      state.dirty = false;

      writeMonthCache(month, {
        students: state.students,
        grid: state.grid,
        dates: [...state.savedDates]
      });

      renderDateChips();
      renderStudentList();
      renderCounters();
    } catch (err) {
      if (!showedCache) {
        toast(err.message, 'error');
        setBrandSubtitle('Falha ao carregar');
      } else {
        toast('Sem conexão; exibindo dados em cache.', 'error');
        renderCounters();
      }
    } finally {
      hideLoading();
    }
  }

  async function saveCurrent() {
    if (!state.date) { toast('Selecione uma data.', 'error'); return; }
    if (state.students.length === 0) { toast('Sem alunos para salvar.', 'error'); return; }

    const attendance = {};
    state.students.forEach((name) => {
      attendance[name] = (state.grid[name] && state.grid[name][state.date]) || '';
    });

    showLoading('Salvando…');
    try {
      await api('saveAttendance', {
        month: state.month,
        date: state.date,
        attendance
      });
      state.dirty = false;
      state.savedDates.add(state.date);
      writeMonthCache(state.month, {
        students: state.students,
        grid: state.grid,
        dates: [...state.savedDates]
      });
      renderDateChips();
      toast('Chamada salva.', 'success');
    } catch (err) {
      toast('Erro ao salvar: ' + err.message, 'error');
    } finally {
      hideLoading();
    }
  }

  async function addStudent(name) {
    const clean = String(name || '').trim();
    if (!clean) { toast('Digite um nome.', 'error'); return false; }
    showLoading('Adicionando…');
    try {
      await api('addStudent', { name: clean });
      toast(`${clean} adicionado.`, 'success');
      await loadMonth(state.month, { preserveDate: true, skipCache: true });
      return true;
    } catch (err) {
      toast('Erro: ' + err.message, 'error');
      return false;
    } finally {
      hideLoading();
    }
  }

  /* ========================================================
   *  Export PNG
   * ======================================================== */

  async function exportPng() {
    if (typeof html2canvas !== 'function') {
      toast('Biblioteca de imagem não carregou.', 'error');
      return;
    }
    if (state.students.length === 0) {
      toast('Sem alunos para exportar.', 'error');
      return;
    }
    buildExportSurface();
    showLoading('Gerando imagem…');
    try {
      const canvas = await html2canvas(el.exportSurface.firstElementChild, {
        backgroundColor: '#ffffff',
        scale: 2,
        useCORS: true,
        logging: false
      });
      const link = document.createElement('a');
      link.download = `chamada-${state.month}.png`;
      link.href = canvas.toDataURL('image/png');
      document.body.appendChild(link);
      link.click();
      link.remove();
      toast('Imagem exportada.', 'success');
    } catch (err) {
      toast('Falha ao gerar imagem.', 'error');
      console.error(err);
    } finally {
      hideLoading();
      el.exportSurface.innerHTML = '';
    }
  }

  function buildExportSurface() {
    const dates = state.sundays.slice();
    const card = document.createElement('div');
    card.className = 'export-card';

    const churchName = (window.APP_CONFIG && window.APP_CONFIG.CHURCH_NAME) || 'EBD';
    const title = document.createElement('h1');
    title.textContent = `${churchName} · Chamada ${formatMonthLong(state.month)}`;
    const sub = document.createElement('p');
    sub.className = 'sub';
    sub.textContent = `Gerado em ${new Date().toLocaleString('pt-BR')}`;
    card.appendChild(title);
    card.appendChild(sub);

    const legend = document.createElement('div');
    legend.className = 'legend';
    STATUSES.forEach((s) => {
      const sp = document.createElement('span');
      sp.style.background = STATUS_COLORS[s];
      sp.textContent = `${s} — ${STATUS_LABELS[s]}`;
      legend.appendChild(sp);
    });
    card.appendChild(legend);

    // Tabela
    const table = document.createElement('table');
    table.style.borderCollapse = 'collapse';
    table.style.width = '100%';
    table.style.fontSize = '13px';

    const thead = document.createElement('thead');
    const trh = document.createElement('tr');
    const thName = document.createElement('th');
    thName.textContent = 'Aluno';
    styleCell(thName, true);
    thName.style.textAlign = 'left';
    trh.appendChild(thName);
    dates.forEach((d) => {
      const th = document.createElement('th');
      th.innerHTML = `Dom<br><b>${formatDateShort(d)}</b>`;
      styleCell(th, true);
      trh.appendChild(th);
    });
    const thPct = document.createElement('th');
    thPct.textContent = '% Pres.';
    styleCell(thPct, true);
    trh.appendChild(thPct);
    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    state.students.forEach((name, idx) => {
      const tr = document.createElement('tr');
      if (idx % 2 === 1) tr.style.background = '#f8fafc';
      const tdName = document.createElement('td');
      tdName.textContent = name;
      styleCell(tdName);
      tdName.style.textAlign = 'left';
      tdName.style.fontWeight = '600';
      tr.appendChild(tdName);

      let presentLike = 0, considered = 0;
      dates.forEach((d) => {
        const s = (state.grid[name] && state.grid[name][d]) || '';
        const td = document.createElement('td');
        styleCell(td);
        if (s) {
          const pill = document.createElement('span');
          pill.textContent = s;
          pill.style.background = STATUS_COLORS[s];
          pill.style.color = '#fff';
          pill.style.fontWeight = '800';
          pill.style.fontSize = '11px';
          pill.style.padding = '2px 7px';
          pill.style.borderRadius = '999px';
          td.appendChild(pill);
          considered++;
          if (s === 'P' || s === 'A') presentLike++;
        } else {
          td.textContent = '—';
          td.style.color = '#94a3b8';
        }
        tr.appendChild(td);
      });

      const tdPct = document.createElement('td');
      tdPct.textContent = considered ? `${Math.round((presentLike / considered) * 100)}%` : '—';
      styleCell(tdPct);
      tdPct.style.fontWeight = '700';
      tr.appendChild(tdPct);

      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    // Resumo do mês
    const summary = document.createElement('div');
    summary.style.marginTop = '14px';
    summary.style.fontSize = '12px';
    summary.style.color = '#475569';
    summary.textContent =
      `Total de alunos: ${state.students.length} · ` +
      `Domingos no mês: ${dates.length}`;

    card.appendChild(table);
    card.appendChild(summary);

    el.exportSurface.innerHTML = '';
    el.exportSurface.appendChild(card);
  }

  function styleCell(td, isHeader) {
    td.style.border = '1px solid #e5e7eb';
    td.style.padding = '8px 10px';
    td.style.textAlign = 'center';
    if (isHeader) {
      td.style.background = '#f1f5f9';
      td.style.fontWeight = '700';
    }
  }

  /* ========================================================
   *  Eventos
   * ======================================================== */

  function bindEvents() {
    el.monthInput.addEventListener('change', () => {
      const val = el.monthInput.value;
      if (!/^\d{4}-\d{2}$/.test(val)) return;
      if (state.dirty && !confirm('Alterações não salvas serão perdidas. Continuar?')) {
        el.monthInput.value = state.month;
        return;
      }
      loadMonth(val);
    });

    el.btnSave.addEventListener('click', saveCurrent);

    el.btnHistory.addEventListener('click', () => {
      renderHistory();
      el.historyModal.showModal();
    });

    el.btnAdd.addEventListener('click', () => {
      el.newName.value = '';
      el.addModal.showModal();
      setTimeout(() => el.newName.focus(), 50);
    });
    el.addForm.addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const ok = await addStudent(el.newName.value);
      if (ok) el.addModal.close();
    });
    el.addModal.addEventListener('click', (ev) => {
      if (ev.target.matches('[data-close]')) el.addModal.close();
    });

    el.btnExport.addEventListener('click', exportPng);

    // Avisa quando o usuário tenta sair com alterações
    window.addEventListener('beforeunload', (ev) => {
      if (state.dirty) {
        ev.preventDefault();
        ev.returnValue = '';
      }
    });
  }

  /* ========================================================
   *  Init
   * ======================================================== */

  function init() {
    try { getApiUrl(); }
    catch (err) {
      setBrandSubtitle('config.js pendente');
      toast(err.message, 'error');
      return;
    }
    bindEvents();
    loadMonth(todayMonth());
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();

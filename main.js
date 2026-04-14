const STORAGE_KEY = "escala-psar-workspace-v2";
const ACTIVE_SESSION_KEY = "escala-psar-session-v1";
const DEFAULT_ACCOUNTS = [];
const FIREBASE_NAMESPACE = "escalaPSARApp";
const FIREBASE_CONFIG = {
  apiKey: "AIzaSyBR0io-r_snZTWy1pe8A0dsb4awBpANDxs",
  authDomain: "escala-pnae.firebaseapp.com",
  databaseURL: "https://escala-pnae-default-rtdb.firebaseio.com",
  projectId: "escala-pnae",
  storageBucket: "escala-pnae.firebasestorage.app",
  messagingSenderId: "289230851590",
  appId: "1:289230851590:web:a638d6c6a8409d65803d87",
};

const SHIFT_OPTIONS = {
  "6h": [
    { value: "H1", label: "H1 16:15-22:15" },
    { value: "H2", label: "H2 16:45-22:45" },
    { value: "H3", label: "H3 17:30-23:30" },
  ],
  "4h": [
    { value: "4A", label: "H4 18:30-22:30" },
    { value: "4B", label: "H5 19:30-23:30" },
  ],
};

const STATUS_LABELS = {
  OFF: "Folga",
  ferias: "Férias",
  atestado: "Atestado",
  compensa: "Folga compensa",
  bloqueio: "Indisponível",
  DISP: "Disponível",
  H1: "H1 16:15-22:15",
  H2: "H2 16:45-22:45",
  H3: "H3 17:30-23:30",
  "4A": "H4 18:30-22:30",
  "4B": "H5 19:30-23:30",
};

const workspace = loadWorkspace();
let state = ensureActiveSchedule();
let cloudRepository = null;
const dom = {};

document.addEventListener("DOMContentLoaded", () => {
  cacheDom();
  hydrateControls();
  bindEvents();
  initializeApp();
});

function cacheDom() {
  dom.loginScreen = document.getElementById("loginScreen");
  dom.loginForm = document.getElementById("loginForm");
  dom.loginUsername = document.getElementById("loginUsername");
  dom.loginPassword = document.getElementById("loginPassword");
  dom.loginFeedback = document.getElementById("loginFeedback");
  dom.appShell = document.getElementById("appShell");
  dom.logoutBtn = document.getElementById("logoutBtn");
  dom.scheduleList = document.getElementById("scheduleList");
  dom.newScheduleForm = document.getElementById("newScheduleForm");
  dom.newScheduleName = document.getElementById("newScheduleName");
  dom.monthInput = document.getElementById("monthInput");
  dom.needH1 = document.getElementById("needH1");
  dom.needH2 = document.getElementById("needH2");
  dom.needH3 = document.getElementById("needH3");
  dom.need4A = document.getElementById("need4A");
  dom.need4B = document.getElementById("need4B");
  dom.generateBtn = document.getElementById("generateBtn");
  dom.clearManualBtn = document.getElementById("clearManualBtn");
  dom.employeeForm = document.getElementById("employeeForm");
  dom.employeeTableBody = document.getElementById("employeeTableBody");
  dom.occurrenceForm = document.getElementById("occurrenceForm");
  dom.occEmployee = document.getElementById("occEmployee");
  dom.occurrenceList = document.getElementById("occurrenceList");
  dom.copySource = document.getElementById("copySource");
  dom.copyTarget = document.getElementById("copyTarget");
  dom.copyScheduleBtn = document.getElementById("copyScheduleBtn");
  dom.copyFeedback = document.getElementById("copyFeedback");
  dom.warningBox = document.getElementById("warningBox");
  dom.scheduleMeta = document.getElementById("scheduleMeta");
  dom.scheduleGrid = document.getElementById("scheduleGrid");
  dom.printSchedule = document.getElementById("printSchedule");
  dom.exportPdfBtn = document.getElementById("exportPdfBtn");
  dom.dayModal = document.getElementById("dayModal");
  dom.closeModalBtn = document.getElementById("closeModalBtn");
  dom.modalTitle = document.getElementById("modalTitle");
  dom.modalSummary = document.getElementById("modalSummary");
  dom.modalPeopleList = document.getElementById("modalPeopleList");
}

function createEmptySchedule(name = "Escala principal") {
  const today = new Date();
  const fallbackMonth = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}`;
  return {
    id: crypto.randomUUID(),
    name,
    settings: { month: fallbackMonth, needH1: 3, needH2: 0, needH3: 1, need4A: 2, need4B: 1 },
    employees: [],
    occurrences: [],
    manualOverrides: {},
    lastGenerated: null,
    updatedAt: new Date().toISOString(),
  };
}

function normalizeSchedule(rawSchedule, fallbackName = "Escala principal") {
  const base = createEmptySchedule(rawSchedule?.name || fallbackName);
  return {
    ...base,
    ...rawSchedule,
    id: rawSchedule?.id || base.id,
    name: rawSchedule?.name || base.name,
    settings: { ...base.settings, ...(rawSchedule?.settings || {}) },
    employees: Array.isArray(rawSchedule?.employees) ? rawSchedule.employees : [],
    occurrences: Array.isArray(rawSchedule?.occurrences) ? rawSchedule.occurrences : [],
    manualOverrides: rawSchedule?.manualOverrides && typeof rawSchedule.manualOverrides === "object" ? rawSchedule.manualOverrides : {},
    lastGenerated: rawSchedule?.lastGenerated || null,
    updatedAt: rawSchedule?.updatedAt || base.updatedAt,
  };
}

function loadWorkspace() {
  const raw = localStorage.getItem(STORAGE_KEY);
  const session = localStorage.getItem(ACTIVE_SESSION_KEY) || "";
  if (!raw) {
    const firstSchedule = createEmptySchedule("Escala principal");
    return {
      accounts: [],
      sessionUser: session || null,
      activeScheduleId: firstSchedule.id,
      schedules: [firstSchedule],
    };
  }

  try {
    const parsed = JSON.parse(raw);
    const schedules =
      Array.isArray(parsed.schedules) && parsed.schedules.length
        ? parsed.schedules.map((schedule, index) => normalizeSchedule(schedule, `Escala ${index + 1}`))
        : [createEmptySchedule("Escala principal")];
    return {
      accounts: parsed.accounts?.length ? parsed.accounts : [],
      sessionUser: session || parsed.sessionUser || null,
      activeScheduleId: parsed.activeScheduleId || schedules[0].id,
      schedules,
    };
  } catch (error) {
    const firstSchedule = createEmptySchedule("Escala principal");
    return {
      accounts: DEFAULT_ACCOUNTS,
      sessionUser: session || null,
      activeScheduleId: firstSchedule.id,
      schedules: [firstSchedule],
    };
  }
}

function ensureActiveSchedule() {
  let schedule = workspace.schedules.find((item) => item.id === workspace.activeScheduleId);
  if (!schedule) {
    schedule = workspace.schedules[0] || createEmptySchedule("Escala principal");
    if (!workspace.schedules.length) workspace.schedules.push(schedule);
    workspace.activeScheduleId = schedule.id;
  }
  return normalizeSchedule(schedule, schedule.name || "Escala principal");
}

function saveState() {
  state.updatedAt = new Date().toISOString();
  const index = workspace.schedules.findIndex((item) => item.id === state.id);
  if (index >= 0) {
    workspace.schedules[index] = structuredClone(state);
  } else {
    workspace.schedules.push(structuredClone(state));
  }
  localStorage.setItem(STORAGE_KEY, JSON.stringify(workspace));
  localStorage.setItem(ACTIVE_SESSION_KEY, workspace.sessionUser || "");
  persistWorkspaceRemote();
}

function hydrateControls() {
  dom.monthInput.value = state.settings.month;
  dom.needH1.value = state.settings.needH1;
  dom.needH2.value = state.settings.needH2;
  dom.needH3.value = state.settings.needH3;
  dom.need4A.value = state.settings.need4A;
  dom.need4B.value = state.settings.need4B;
}

function bindEvents() {
  dom.loginForm.addEventListener("submit", handleLogin);
  dom.logoutBtn.addEventListener("click", logout);
  dom.newScheduleForm.addEventListener("submit", createScheduleFromForm);
  dom.generateBtn.addEventListener("click", generateSchedule);
  dom.exportPdfBtn.addEventListener("click", () => window.print());
  dom.clearManualBtn.addEventListener("click", () => {
    state.manualOverrides = {};
    saveState();
    renderSchedule();
  });
  dom.closeModalBtn.addEventListener("click", closeDayModal);
  dom.dayModal.addEventListener("click", (event) => {
    if (event.target === dom.dayModal) closeDayModal();
  });

  dom.monthInput.addEventListener("change", syncSettingsFromInputs);
  [dom.needH1, dom.needH2, dom.needH3, dom.need4A, dom.need4B].forEach((input) =>
    input.addEventListener("change", syncSettingsFromInputs)
  );

  dom.employeeForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const formData = new FormData(dom.employeeForm);
    state.employees.push({
      id: crypto.randomUUID(),
      registration: formData.get("registration").toString().trim(),
      name: formData.get("name").toString().trim(),
      workload: formData.get("workload").toString(),
      cycle: formData.get("cycle").toString(),
      offDate: formData.get("offDate").toString(),
      firstOffLength: Number(formData.get("firstOffLength")),
      preferences: {
        H1: Number(formData.get("prefH1")) || 0,
        H2: Number(formData.get("prefH2")) || 0,
        H3: Number(formData.get("prefH3")) || 0,
        "4A": Number(formData.get("prefH4")) || 0,
        "4B": Number(formData.get("prefH5")) || 0,
      },
    });
    dom.employeeForm.reset();
    syncSettingsFromInputs();
    renderAll();
  });

  dom.occurrenceForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const formData = new FormData(dom.occurrenceForm);
    if (!formData.get("employeeId")) return;
    state.occurrences.push({
      id: crypto.randomUUID(),
      employeeId: formData.get("employeeId").toString(),
      date: formData.get("date").toString(),
      type: formData.get("type").toString(),
    });
    saveState();
    dom.occurrenceForm.reset();
    renderAll();
  });

  dom.copyScheduleBtn.addEventListener("click", copyMonthlySchedule);
}

async function initializeApp() {
  cloudRepository = await createCloudRepository();
  await hydrateWorkspaceFromCloud();
  if (workspace.sessionUser) {
    dom.loginScreen.setAttribute("hidden", "");
    dom.appShell.removeAttribute("hidden");
    renderAll();
  } else {
    dom.appShell.setAttribute("hidden", "");
    dom.loginScreen.removeAttribute("hidden");
  }
}

function syncSettingsFromInputs() {
  state.settings.month = dom.monthInput.value;
  state.settings.needH1 = Number(dom.needH1.value);
  state.settings.needH2 = Number(dom.needH2.value);
  state.settings.needH3 = Number(dom.needH3.value);
  state.settings.need4A = Number(dom.need4A.value);
  state.settings.need4B = Number(dom.need4B.value);
  saveState();
}

function renderAll() {
  renderScheduleMenu();
  renderEmployees();
  renderOccurrenceFormOptions();
  renderCopyOptions();
  renderOccurrences();
  renderSchedule();
}


function renderScheduleMenu() {
  dom.scheduleList.innerHTML = workspace.schedules
    .map(
      (schedule) => `
        <button type="button" class="schedule-item ${schedule.id === state.id ? "active" : ""}" data-schedule-id="${schedule.id}">
          <strong>${escapeHtml(schedule.name)}</strong>
          <small>${formatMonthYear(schedule.settings.month)}</small>
        </button>
      `
    )
    .join("");

  dom.scheduleList.querySelectorAll("[data-schedule-id]").forEach((button) => {
    button.addEventListener("click", () => switchSchedule(button.dataset.scheduleId));
  });
}

function renderEmployees() {
  if (!state.employees.length) {
    dom.employeeTableBody.innerHTML = `
      <tr>
        <td colspan="7">Nenhuma pessoa cadastrada ainda.</td>
      </tr>
    `;
    return;
  }

  dom.employeeTableBody.innerHTML = state.employees
    .map(
      (employee) => `
        <tr>
          <td><input data-field="registration" data-id="${employee.id}" value="${escapeHtml(employee.registration)}" /></td>
          <td><input data-field="name" data-id="${employee.id}" value="${escapeHtml(employee.name)}" /></td>
          <td>
            <select data-field="workload" data-id="${employee.id}">
              <option value="6h" ${employee.workload === "6h" ? "selected" : ""}>6h</option>
              <option value="4h" ${employee.workload === "4h" ? "selected" : ""}>4h</option>
            </select>
          </td>
          <td><input data-field="offDate" data-id="${employee.id}" type="date" value="${employee.offDate}" /></td>
          <td>
            <select data-field="firstOffLength" data-id="${employee.id}">
              <option value="1" ${Number(employee.firstOffLength) === 1 ? "selected" : ""}>1 dia</option>
              <option value="2" ${Number(employee.firstOffLength) === 2 ? "selected" : ""}>2 dias</option>
            </select>
          </td>
          <td>
            <div class="preferences-editor">
              ${renderPreferenceEditor(employee)}
            </div>
          </td>
          <td><button class="remove-btn" data-remove-employee="${employee.id}">Remover</button></td>
        </tr>
      `
    )
    .join("");

  dom.employeeTableBody.querySelectorAll("input[data-field], select[data-field]").forEach((element) => {
    element.addEventListener("change", handleEmployeeFieldChange);
  });

  dom.employeeTableBody.querySelectorAll("input[data-pref], select[data-pref-workload]").forEach((element) => {
    element.addEventListener("change", handleEmployeePreferenceChange);
  });

  dom.employeeTableBody.querySelectorAll("[data-remove-employee]").forEach((button) => {
    button.addEventListener("click", () => removeEmployee(button.dataset.removeEmployee));
  });
}

function handleEmployeeFieldChange(event) {
  const { id, field } = event.target.dataset;
  const employee = state.employees.find((item) => item.id === id);
  if (!employee) return;
  employee[field] = ["firstOffLength"].includes(field) ? Number(event.target.value) : event.target.value;
  saveState();
  renderAll();
}

function handleEmployeePreferenceChange(event) {
  const employee = state.employees.find((item) => item.id === event.target.dataset.id);
  if (!employee) return;

  if (event.target.dataset.prefWorkload) {
    const workload = event.target.value;
    employee.workload = workload;
    employee.preferences = workload === "6h" ? { H1: 0, H2: 0, H3: 0, "4A": 0, "4B": 0 } : { H1: 0, H2: 0, H3: 0, "4A": 0, "4B": 0 };
    saveState();
    renderAll();
    return;
  }

  const shift = event.target.dataset.pref;
  employee.preferences = employee.preferences || { H1: 0, H2: 0, H3: 0, "4A": 0, "4B": 0 };
  employee.preferences[shift] = Number(event.target.value) || 0;
  saveState();
  renderAll();
}

function removeEmployee(employeeId) {
  state.employees = state.employees.filter((employee) => employee.id !== employeeId);
  state.occurrences = state.occurrences.filter((occurrence) => occurrence.employeeId !== employeeId);
  Object.keys(state.manualOverrides).forEach((key) => {
    if (key.startsWith(`${employeeId}|`)) delete state.manualOverrides[key];
  });
  saveState();
  renderAll();
}

function renderOccurrenceFormOptions() {
  if (!state.employees.length) {
    dom.occEmployee.innerHTML = `<option value="">Cadastre uma pessoa primeiro</option>`;
    return;
  }
  dom.occEmployee.innerHTML = state.employees
    .map(
      (employee) =>
        `<option value="${employee.id}">${escapeHtml(employee.name)} - ${escapeHtml(employee.registration)}</option>`
    )
    .join("");
}

function renderCopyOptions() {
  if (!state.employees.length) {
    dom.copySource.innerHTML = `<option value="">Cadastre uma pessoa primeiro</option>`;
    dom.copyTarget.innerHTML = `<option value="">Cadastre uma pessoa primeiro</option>`;
    dom.copyFeedback.textContent = "";
    dom.copyFeedback.className = "copy-feedback";
    return;
  }

  const options = state.employees
    .map((employee) => `<option value="${employee.id}">${escapeHtml(employee.name)} - ${escapeHtml(employee.registration)}</option>`)
    .join("");

  dom.copySource.innerHTML = options;
  dom.copyTarget.innerHTML = options;
}

function renderOccurrences() {
  if (!state.occurrences.length) {
    dom.occurrenceList.innerHTML = `<div class="chip">Nenhuma ocorrência cadastrada.</div>`;
    return;
  }

  dom.occurrenceList.innerHTML = state.occurrences
    .sort((a, b) => a.date.localeCompare(b.date))
    .map((occurrence) => {
      const employee = state.employees.find((item) => item.id === occurrence.employeeId);
      const employeeName = employee ? employee.name : "Pessoa removida";
      return `
        <div class="chip">
          <span>${formatDateBr(occurrence.date)} - ${escapeHtml(employeeName)} - ${STATUS_LABELS[occurrence.type]}</span>
          <button data-remove-occurrence="${occurrence.id}">x</button>
        </div>
      `;
    })
    .join("");

  dom.occurrenceList.querySelectorAll("[data-remove-occurrence]").forEach((button) => {
    button.addEventListener("click", () => {
      state.occurrences = state.occurrences.filter((occurrence) => occurrence.id !== button.dataset.removeOccurrence);
      saveState();
      renderAll();
    });
  });
}

function copyMonthlySchedule() {
  const sourceId = dom.copySource.value;
  const targetId = dom.copyTarget.value;
  const month = state.settings.month;

  if (!sourceId || !targetId) {
    return setCopyFeedback("Selecione a pessoa de origem e a de destino.", true);
  }

  if (sourceId === targetId) {
    return setCopyFeedback("Escolha pessoas diferentes para copiar a escala.", true);
  }

  const source = state.employees.find((employee) => employee.id === sourceId);
  const target = state.employees.find((employee) => employee.id === targetId);

  if (!source || !target) {
    return setCopyFeedback("Não foi possível localizar as pessoas selecionadas.", true);
  }

  if (source.workload !== target.workload) {
    return setCopyFeedback("Só dá para copiar a escala entre pessoas com a mesma carga horária.", true);
  }

  const result = buildSchedule(state);
  const dates = getMonthDates(month);
  let copied = 0;

  dates.forEach((date) => {
    const sourceValue = result.schedule[date]?.[sourceId];
    if (!sourceValue) return;

    const allowedValues = new Set([
      ...SHIFT_OPTIONS[target.workload].map((option) => option.value),
      "DISP",
      "OFF",
      "ferias",
      "atestado",
      "compensa",
      "bloqueio",
    ]);

    if (!allowedValues.has(sourceValue)) return;
    state.manualOverrides[`${targetId}|${date}`] = sourceValue;
    copied += 1;
  });

  saveState();
  renderSchedule();
  setCopyFeedback(`Escala de ${source.name} copiada para ${target.name} em ${copied} dias do mês.`, false);
}

function setCopyFeedback(message, isError) {
  dom.copyFeedback.textContent = message;
  dom.copyFeedback.className = isError ? "copy-feedback error" : "copy-feedback";
}

function handleLogin(event) {
  event.preventDefault();
  const username = dom.loginUsername.value.trim();
  const password = dom.loginPassword.value.trim();
  const account = workspace.accounts.find((item) => item.username === username && item.password === password);

  if (!account) {
    dom.loginFeedback.textContent = "Usuário ou senha inválidos.";
    return;
  }

  workspace.sessionUser = account.username;
  localStorage.setItem(ACTIVE_SESSION_KEY, workspace.sessionUser);
  dom.loginFeedback.textContent = "";
  dom.loginForm.reset();
  dom.loginScreen.setAttribute("hidden", "");
  dom.appShell.removeAttribute("hidden");
  renderAll();
}

function logout() {
  workspace.sessionUser = null;
  localStorage.removeItem(ACTIVE_SESSION_KEY);
  dom.appShell.setAttribute("hidden", "");
  dom.loginScreen.removeAttribute("hidden");
}

function createScheduleFromForm(event) {
  event.preventDefault();
  const name = dom.newScheduleName.value.trim();
  if (!name) return;
  const schedule = createEmptySchedule(name);
  workspace.schedules.push(schedule);
  workspace.activeScheduleId = schedule.id;
  state = schedule;
  dom.newScheduleForm.reset();
  saveState();
  renderAll();
}

function switchSchedule(scheduleId) {
  const schedule = workspace.schedules.find((item) => item.id === scheduleId);
  if (!schedule) return;
  workspace.activeScheduleId = scheduleId;
  state = structuredClone(schedule);
  saveState();
  hydrateControls();
  renderAll();
}

async function createCloudRepository() {
  if (!window.firebase?.initializeApp || !FIREBASE_CONFIG.apiKey || !FIREBASE_CONFIG.databaseURL) return null;

  try {
    const app = window.firebase.apps.length
      ? window.firebase.app()
      : window.firebase.initializeApp(FIREBASE_CONFIG);

    const db = window.firebase.database(app);
    const workspaceRef = db.ref(`${FIREBASE_NAMESPACE}/workspace`);

    return {
      async load() {
        const snapshot = await workspaceRef.get();
        return snapshot.exists() ? snapshot.val() : null;
      },
      async save(data) {
        await workspaceRef.set(data);
      },
    };
  } catch (error) {
    console.warn("Firebase indisponível. Usando armazenamento local.", error);
    return null;
  }
}

async function hydrateWorkspaceFromCloud() {
  if (!cloudRepository) return;
  try {
    const cloudData = await cloudRepository.load();
    if (!cloudData?.schedules?.length) return;
    workspace.accounts = cloudData.accounts?.length ? cloudData.accounts : workspace.accounts;
    workspace.schedules = cloudData.schedules.map((schedule, index) => normalizeSchedule(schedule, `Escala ${index + 1}`));
    workspace.activeScheduleId = cloudData.activeScheduleId || cloudData.schedules[0].id;
    state = ensureActiveSchedule();
    hydrateControls();
  } catch (error) {
    console.warn("Falha ao carregar do Firebase.", error);
  }
}

async function persistWorkspaceRemote() {
  if (!cloudRepository) return;
  try {
    await cloudRepository.save(workspace);
  } catch (error) {
    console.warn("Falha ao salvar no Firebase.", error);
  }
}

function generateSchedule() {
  syncSettingsFromInputs();
  state.lastGenerated = new Date().toISOString();
  saveState();
  renderSchedule();
}

function renderSchedule() {
  const month = state.settings.month;
  if (!month || !state.employees.length) {
    dom.warningBox.innerHTML = `<div class="warning">Cadastre a equipe e selecione o mês para visualizar a escala.</div>`;
    dom.scheduleMeta.innerHTML = "";
    dom.scheduleGrid.innerHTML = "";
    return;
  }

  const result = buildSchedule(state);
  renderWarnings(result.warnings);
  renderMeta(result);
  renderScheduleTable(result);
  renderPrintSchedule(result);
}

function renderWarnings(warnings) {
  if (!warnings.length) {
    dom.warningBox.innerHTML = `<div class="warning ok">Cobertura mínima atendida em todos os dias, considerando as regras configuradas.</div>`;
    return;
  }

  const uniqueWarnings = [...new Set(warnings)];
  dom.warningBox.innerHTML = `
    <div class="warning warning-summary">
      <div>
        <strong>Há dias com cobertura incompleta na escala.</strong>
        <p>Revise os horários destacados na grade. Alguns dias estão sem a quantidade mínima de pessoas em um ou mais turnos.</p>
      </div>
      <button type="button" class="warning-toggle" id="warningToggleBtn">Saiba mais (${uniqueWarnings.length})</button>
    </div>
    <div class="warning-details" id="warningDetails" hidden>
      ${uniqueWarnings.map((warning) => `<div class="warning-detail-item">${escapeHtml(warning)}</div>`).join("")}
    </div>
  `;

  const toggleButton = document.getElementById("warningToggleBtn");
  const detailsBox = document.getElementById("warningDetails");
  toggleButton.addEventListener("click", () => {
    const isHidden = detailsBox.hasAttribute("hidden");
    if (isHidden) {
      detailsBox.removeAttribute("hidden");
      toggleButton.textContent = "Ocultar detalhes";
    } else {
      detailsBox.setAttribute("hidden", "");
      toggleButton.textContent = `Saiba mais (${uniqueWarnings.length})`;
    }
  });
}

function renderMeta(result) {
  const count6h = state.employees.filter((employee) => employee.workload === "6h").length;
  const count4h = state.employees.filter((employee) => employee.workload === "4h").length;
  const manualCount = Object.keys(state.manualOverrides).length;
  dom.scheduleMeta.innerHTML = `
    <div class="meta-pill">Equipe 6h: ${count6h}</div>
    <div class="meta-pill">Equipe 4h: ${count4h}</div>
    <div class="meta-pill">Dias no mês: ${result.dates.length}</div>
    <div class="meta-pill">Ajustes manuais: ${manualCount}</div>
    <div class="meta-pill">Última geração: ${state.lastGenerated ? new Date(state.lastGenerated).toLocaleString("pt-BR") : "ainda não gerada"}</div>
  `;
}

function renderScheduleTable(result) {
  const weekdayNames = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sab"];
  const firstWeekday = getWeekdayIndex(result.dates[0]);
  const emptyCells = Array.from({ length: firstWeekday }, () => `<div class="calendar-empty"></div>`).join("");

  const dayCards = result.dates
    .map((date) => {
      const coverage = result.coverageByDate[date];
      const isAlert = coverage.missing.length > 0;
      const assignedPeople = state.employees
        .map((employee) => ({ employee, value: result.schedule[date][employee.id] }))
        .filter((item) => ["H1", "H2", "H3", "4A", "4B"].includes(item.value))
        .slice(0, 4);

      const peoplePreview = assignedPeople
        .map(
          (item) =>
            `<div class="calendar-person-chip">${escapeHtml(item.employee.name)} · ${escapeHtml(STATUS_LABELS[item.value])}</div>`
        )
        .join("");
      const moreCount =
        Object.values(result.schedule[date]).filter((value) => ["H1", "H2", "H3", "4A", "4B"].includes(value)).length -
        assignedPeople.length;

      return `
        <div class="calendar-day ${isAlert ? "is-alert" : ""}">
          <button type="button" data-open-day="${date}">
            <div class="calendar-day-header">
              <div>
                <span class="calendar-day-number">${date.slice(-2)}</span>
                <span class="calendar-day-week">${getWeekdayLabel(date)}</span>
              </div>
              <span class="calendar-status ${isAlert ? "warn" : "ok"}">${isAlert ? "Ajustar" : "Ok"}</span>
            </div>
            <div class="calendar-coverage">
              <div class="calendar-coverage-line"><strong>6h</strong> H1 ${coverage.counts.H1} · H2 ${coverage.counts.H2} · H3 ${coverage.counts.H3}</div>
              <div class="calendar-coverage-line"><strong>4h</strong> H4 ${coverage.counts["4A"]} · H5 ${coverage.counts["4B"]}</div>
            </div>
            <div class="calendar-people-preview">
              ${peoplePreview || `<div class="calendar-person-chip more">Nenhuma pessoa alocada</div>`}
              ${moreCount > 0 ? `<div class="calendar-person-chip more">+${moreCount} pessoas</div>` : ""}
            </div>
          </button>
        </div>
      `;
    })
    .join("");

  dom.scheduleGrid.innerHTML = `
    <div class="calendar-shell">
      <div class="calendar-weekdays">
        ${weekdayNames.map((name) => `<div class="calendar-weekday">${name}</div>`).join("")}
      </div>
      <div class="calendar-grid">
        ${emptyCells}
        ${dayCards}
      </div>
    </div>
  `;

  dom.scheduleGrid.querySelectorAll("[data-open-day]").forEach((button) => {
    button.addEventListener("click", () => openDayModal(button.dataset.openDay, result));
  });
}

function renderPrintSchedule(result) {
  const dayHeaders = result.dates
    .map((date) => {
      const [, , day] = date.split("-");
      return `
        <th class="print-day-col">
          <span>${day}</span>
          <small>${getWeekdayLabel(date)}</small>
        </th>
      `;
    })
    .join("");

  const rows = state.employees
    .map((employee) => {
      const dayCells = result.dates
        .map((date) => {
          const value = result.schedule[date][employee.id];
          return `<td class="print-shift-cell ${getPrintCellClass(value)}">${escapeHtml(getCompactLabel(value))}</td>`;
        })
        .join("");

      return `
        <tr>
          <td class="print-registration">${escapeHtml(employee.registration)}</td>
          <td class="print-name">${escapeHtml(employee.name)}</td>
          ${dayCells}
        </tr>
      `;
    })
    .join("");

  dom.printSchedule.innerHTML = `
    <div class="print-sheet">
      <div class="print-sheet-header">
        <div>
          <strong>Escala PSAR</strong>
          <span>${formatMonthYear(state.settings.month)}</span>
        </div>
      </div>
      <div class="print-legend-panel">
        <div class="print-legend-title">Legenda de horários</div>
        <div class="print-sheet-legend">
          <span><strong>H1</strong> 16:15-22:15</span>
          <span><strong>H2</strong> 16:45-22:45</span>
          <span><strong>H3</strong> 17:30-23:30</span>
          <span><strong>H4</strong> 18:30-22:30</span>
          <span><strong>H5</strong> 19:30-23:30</span>
          <span><strong>F</strong> Folga</span>
          <span><strong>FC</strong> Folga compensa</span>
          <span><strong>V</strong> Férias</span>
          <span><strong>DM</strong> Atestado</span>
        </div>
      </div>
      <div class="print-table-wrap">
        <table class="print-table">
          <thead>
            <tr>
              <th class="print-registration">Matrícula</th>
              <th class="print-name">Nome</th>
              ${dayHeaders}
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
    </div>
  `;
}

function handleCellChange(event) {
  const date = event.target.dataset.date;
  const employeeId = event.target.dataset.employee;
  const key = `${employeeId}|${date}`;
  if (event.target.value === "AUTO") {
    delete state.manualOverrides[key];
  } else {
    state.manualOverrides[key] = event.target.value;
  }
  saveState();
  renderSchedule();
  if (!dom.dayModal.hasAttribute("hidden")) {
    openDayModal(date, buildSchedule(state));
  }
}

function openDayModal(date, result) {
  dom.modalTitle.textContent = `Dia ${formatDateLong(date)}`;
  const coverage = result.coverageByDate[date];
  dom.modalSummary.innerHTML = `
    <span class="meta-pill">H1 ${coverage.counts.H1}</span>
    <span class="meta-pill">H2 ${coverage.counts.H2}</span>
    <span class="meta-pill">H3 ${coverage.counts.H3}</span>
    <span class="meta-pill">H4 ${coverage.counts["4A"]}</span>
    <span class="meta-pill">H5 ${coverage.counts["4B"]}</span>
  `;

  dom.modalPeopleList.innerHTML = state.employees
    .map((employee) => renderModalPersonRow(employee, date, result.schedule[date][employee.id], result.autoSchedule[date][employee.id]))
    .join("");

  dom.modalPeopleList.querySelectorAll(".cell-select").forEach((select) => {
    select.addEventListener("change", handleCellChange);
  });

  dom.dayModal.removeAttribute("hidden");
}

function closeDayModal() {
  dom.dayModal.setAttribute("hidden", "");
}

function renderModalPersonRow(employee, date, currentValue, autoValue) {
  const options = [...SHIFT_OPTIONS[employee.workload], { value: "DISP", label: STATUS_LABELS.DISP }];
  const statusOptions = ["OFF", "ferias", "atestado", "compensa", "bloqueio"];
  const selectedValue = currentValue === autoValue ? "AUTO" : currentValue;
  const optionMarkup = [
    `<option value="AUTO" ${selectedValue === "AUTO" ? "selected" : ""}>Auto (${STATUS_LABELS[autoValue] || autoValue})</option>`,
    ...options.map(
      (option) => `<option value="${option.value}" ${selectedValue === option.value ? "selected" : ""}>${option.label}</option>`
    ),
    ...statusOptions.map(
      (status) => `<option value="${status}" ${selectedValue === status ? "selected" : ""}>${STATUS_LABELS[status]}</option>`
    ),
  ].join("");

  return `
    <div class="modal-person-row">
      <div class="modal-person-name">
        <strong>${escapeHtml(employee.name)}</strong>
        <small>${escapeHtml(employee.registration)} · ${employee.workload}</small>
      </div>
      <div class="modal-person-status">
        Atual: <span class="status-badge ${isWorkShift(currentValue) ? "status-ok" : "status-warn"}">${escapeHtml(
    STATUS_LABELS[currentValue] || currentValue
  )}</span>
      </div>
      <select class="cell-select" data-date="${date}" data-employee="${employee.id}">
        ${optionMarkup}
      </select>
    </div>
  `;
}

function buildSchedule(appState) {
  const dates = getMonthDates(appState.settings.month);
  const occurrencesByEmployeeDate = indexOccurrences(appState.occurrences);
  const autoSchedule = {};
  const finalSchedule = {};
  const warnings = [];
  const coverageByDate = {};
  const stats = {};

  appState.employees.forEach((employee) => {
    stats[employee.id] = { total: 0, H1: 0, H2: 0, H3: 0, "4A": 0, "4B": 0 };
  });

  for (const date of dates) {
    autoSchedule[date] = {};
    finalSchedule[date] = {};
    const available6h = [];
    const available4h = [];

    appState.employees.forEach((employee) => {
      const baseStatus = getBaseStatus(employee, date, occurrencesByEmployeeDate);
      autoSchedule[date][employee.id] = baseStatus;
      finalSchedule[date][employee.id] = baseStatus;
      if (baseStatus === "DISP") {
        if (employee.workload === "6h") {
          available6h.push(employee);
        } else {
          available4h.push(employee);
        }
      }
    });

    applyManualPreAssignments(appState, date, finalSchedule[date], available6h, available4h, stats);

    const missing = [];
    assignMandatoryShift(finalSchedule[date], autoSchedule[date], available6h, "H1", appState.settings.needH1, stats, missing, date);
    assignMandatoryShift(finalSchedule[date], autoSchedule[date], available6h, "H3", appState.settings.needH3, stats, missing, date);

    assignPreferred4hMix(finalSchedule[date], autoSchedule[date], available4h, stats);
    assignRemaining6hMix(finalSchedule[date], autoSchedule[date], available6h, stats);

    applyManualFinalOverrides(appState, date, finalSchedule[date]);

    const coverage = summarizeCoverage(finalSchedule[date], appState.settings);
    coverageByDate[date] = coverage;
    warnings.push(...coverage.missing.map((item) => `${formatShortDate(date)} - ${item}`));
    warnings.push(...missing);
  }

  return { dates, autoSchedule, schedule: finalSchedule, warnings: [...new Set(warnings)], coverageByDate };
}

function applyManualPreAssignments(appState, date, finalDay, pool6h, pool4h, stats) {
  appState.employees.forEach((employee) => {
    const key = `${employee.id}|${date}`;
    const manualValue = appState.manualOverrides[key];
    if (!manualValue) return;

    const allowedSlots = new Set([
      ...SHIFT_OPTIONS[employee.workload].map((option) => option.value),
      "DISP",
      "OFF",
      "ferias",
      "atestado",
      "compensa",
      "bloqueio",
    ]);
    if (!allowedSlots.has(manualValue)) return;

    finalDay[employee.id] = manualValue;
    removeFromPool(employee.workload === "6h" ? pool6h : pool4h, employee.id);

    if (["H1", "H2", "H3", "4A", "4B"].includes(manualValue)) {
      stats[employee.id].total += 1;
      stats[employee.id][manualValue] += 1;
    }
  });
}

function applyManualFinalOverrides(appState, date, finalDay) {
  appState.employees.forEach((employee) => {
    const key = `${employee.id}|${date}`;
    const manualValue = appState.manualOverrides[key];
    if (!manualValue) return;

    const allowedSlots = new Set([
      ...SHIFT_OPTIONS[employee.workload].map((option) => option.value),
      "DISP",
      "OFF",
      "ferias",
      "atestado",
      "compensa",
      "bloqueio",
    ]);
    if (allowedSlots.has(manualValue)) {
      finalDay[employee.id] = manualValue;
    }
  });
}

function assignMandatoryShift(finalDay, autoDay, pool, shift, count, stats, missing, date) {
  const alreadyAssigned = Object.values(finalDay).filter((value) => value === shift).length;
  let needed = Math.max(0, count - alreadyAssigned);
  while (needed > 0) {
    const candidate = pickWeightedCandidate(pool, shift, stats);
    if (!candidate) {
      missing.push(`${formatShortDate(date)} sem cobertura suficiente para ${STATUS_LABELS[shift]}.`);
      break;
    }
    assignShiftToEmployee(finalDay, autoDay, candidate, shift, stats, pool);
    needed -= 1;
  }
}

function assignPreferred4hMix(finalDay, autoDay, pool, stats) {
  if (!pool.length) return;

  const hasH5Manual = Object.values(finalDay).includes("4B");
  if (!hasH5Manual) {
    const firstH5 = pickWeightedCandidate(pool, "4B", stats, 1.15);
    if (firstH5) assignShiftToEmployee(finalDay, autoDay, firstH5, "4B", stats, pool);
  }

  while (pool.length) {
    const shift = chooseShiftForEmployee(pool[0], ["4A", "4B"], stats, { defaultShift: "4A", randomness: 0.45 });
    const candidate = pickWeightedCandidate(pool, shift, stats);
    if (!candidate) break;
    const chosenShift = chooseShiftForEmployee(candidate, ["4A", "4B"], stats, { defaultShift: "4A", randomness: 0.45 });
    assignShiftToEmployee(finalDay, autoDay, candidate, chosenShift, stats, pool);
  }
}

function assignRemaining6hMix(finalDay, autoDay, pool, stats) {
  while (pool.length) {
    const candidate = pickWeightedCandidate(pool, "H2", stats, 1.05) || pool[0];
    const chosenShift = chooseShiftForEmployee(candidate, ["H2", "H3"], stats, { defaultShift: "H2", randomness: 0.3 });
    assignShiftToEmployee(finalDay, autoDay, candidate, chosenShift, stats, pool);
  }
}

function pickWeightedCandidate(pool, targetShift, stats, targetBias = 1) {
  if (!pool.length) return null;
  const weighted = pool.map((employee) => {
    const preferenceBoost = 1 + (getShiftPreference(employee, targetShift) / 100) * targetBias;
    const fairnessBoost = 1 / (1 + stats[employee.id][targetShift] + stats[employee.id].total * 0.08);
    const randomBoost = 0.8 + Math.random() * 0.4;
    return { employee, score: preferenceBoost * fairnessBoost * randomBoost };
  });

  weighted.sort((a, b) => b.score - a.score);
  const topSlice = weighted.slice(0, Math.min(3, weighted.length));
  return topSlice[Math.floor(Math.random() * topSlice.length)].employee;
}

function chooseShiftForEmployee(employee, shifts, stats, options = {}) {
  const validShifts = shifts.filter((shift) => isShiftAllowedForEmployee(employee, shift));
  if (!validShifts.length) return shifts[0];

  const randomness = options.randomness ?? 0.35;
  const weighted = validShifts.map((shift) => {
    const preferenceBoost = 1 + getShiftPreference(employee, shift) / 100;
    const fairnessBoost = 1 / (1 + stats[employee.id][shift] * 0.8);
    const defaultBoost = options.defaultShift === shift ? 1.18 : 1;
    const randomBoost = 1 - randomness / 2 + Math.random() * randomness;
    return { shift, score: preferenceBoost * fairnessBoost * defaultBoost * randomBoost };
  });

  weighted.sort((a, b) => b.score - a.score);
  return weighted[0].shift;
}

function isShiftAllowedForEmployee(employee, shift) {
  return SHIFT_OPTIONS[employee.workload].some((option) => option.value === shift);
}

function getShiftPreference(employee, shift) {
  return Number(employee.preferences?.[shift]) || 0;
}

function assignShiftToEmployee(finalDay, autoDay, employee, shift, stats, pool) {
  finalDay[employee.id] = shift;
  autoDay[employee.id] = shift;
  stats[employee.id].total += 1;
  stats[employee.id][shift] += 1;
  removeFromPool(pool, employee.id);
}

function removeFromPool(pool, employeeId) {
  const index = pool.findIndex((employee) => employee.id === employeeId);
  if (index >= 0) pool.splice(index, 1);
}

function summarizeCoverage(daySchedule, settings) {
  const counts = { H1: 0, H2: 0, H3: 0, "4A": 0, "4B": 0 };
  Object.values(daySchedule).forEach((value) => {
    if (counts[value] !== undefined) counts[value] += 1;
  });

  const missing = [];
  if (counts.H1 < settings.needH1) missing.push(`faltam ${settings.needH1 - counts.H1} pessoas em H1`);
  if (counts.H3 < settings.needH3) missing.push(`faltam ${settings.needH3 - counts.H3} pessoas em H3`);
  if (settings.need4B > 0 && counts["4B"] === 0 && counts["4A"] > 0) missing.push(`nenhuma pessoa ficou em H5`);

  return {
    summary: buildCoverageSummary(counts),
    counts,
    missing,
  };
}

function buildCoverageSummary(counts) {
  return ["H1", counts.H1, "H2", counts.H2, "H3", counts.H3, "H4", counts["4A"], "H5", counts["4B"]].join(" ");
}

function indexOccurrences(occurrences) {
  const map = new Map();
  occurrences.forEach((occurrence) => {
    map.set(`${occurrence.employeeId}|${occurrence.date}`, occurrence.type);
  });
  return map;
}

function getBaseStatus(employee, date, occurrencesByEmployeeDate) {
  const explicit = occurrencesByEmployeeDate.get(`${employee.id}|${date}`);
  if (explicit) return explicit;
  return isEmployeeOff(employee, date) ? "OFF" : "DISP";
}

function isEmployeeOff(employee, targetDate) {
  if (!employee.offDate) return false;
  let currentStart = employee.offDate;
  let currentLength = Number(employee.firstOffLength) || 1;

  if (targetDate >= currentStart) {
    while (currentStart <= targetDate) {
      if (dateInBlock(targetDate, currentStart, currentLength)) return true;
      currentStart = shiftDate(currentStart, currentLength + 6);
      currentLength = currentLength === 1 ? 2 : 1;
    }
    return false;
  }

  while (currentStart > targetDate) {
    const previousLength = currentLength === 1 ? 2 : 1;
    currentStart = shiftDate(currentStart, -(previousLength + 6));
    currentLength = previousLength;
    if (dateInBlock(targetDate, currentStart, currentLength)) return true;
  }
  return false;
}

function dateInBlock(targetDate, blockStart, blockLength) {
  const lastDay = shiftDate(blockStart, blockLength - 1);
  return targetDate >= blockStart && targetDate <= lastDay;
}

function getMonthDates(monthValue) {
  const [year, month] = monthValue.split("-").map(Number);
  const totalDays = new Date(year, month, 0).getDate();
  const dates = [];
  for (let day = 1; day <= totalDays; day += 1) {
    dates.push(`${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`);
  }
  return dates;
}

function shiftDate(isoDate, amount) {
  const [year, month, day] = isoDate.split("-").map(Number);
  const date = new Date(year, month - 1, day);
  date.setDate(date.getDate() + amount);
  return formatDateIso(date);
}

function formatDateIso(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatDateBr(isoDate) {
  const [year, month, day] = isoDate.split("-");
  return `${day}/${month}/${year}`;
}

function formatShortDate(isoDate) {
  const [year, month, day] = isoDate.split("-").map(Number);
  const date = new Date(year, month - 1, day);
  const weekday = date.toLocaleDateString("pt-BR", { weekday: "short" }).replace(".", "");
  return `${String(day).padStart(2, "0")}/${String(month).padStart(2, "0")} ${weekday}`;
}

function formatDateLong(isoDate) {
  const [year, month, day] = isoDate.split("-").map(Number);
  return new Date(year, month - 1, day).toLocaleDateString("pt-BR", {
    day: "2-digit",
    month: "long",
    year: "numeric",
    weekday: "long",
  });
}

function getWeekdayLabel(isoDate) {
  const [year, month, day] = isoDate.split("-").map(Number);
  return new Date(year, month - 1, day).toLocaleDateString("pt-BR", { weekday: "short" }).replace(".", "");
}

function getWeekdayIndex(isoDate) {
  const [year, month, day] = isoDate.split("-").map(Number);
  return new Date(year, month - 1, day).getDay();
}

function isWorkShift(value) {
  return ["H1", "H2", "H3", "4A", "4B"].includes(value);
}

function getCompactLabel(value) {
  const compact = {
    H1: "H1",
    H2: "H2",
    H3: "H3",
    "4A": "H4",
    "4B": "H5",
    OFF: "F",
    ferias: "V",
    atestado: "DM",
    compensa: "FC",
    bloqueio: "IND",
    DISP: "-",
  };
  return compact[value] || value;
}

function getPrintCellClass(value) {
  if (value === "OFF") return "is-off";
  if (["ferias", "atestado", "compensa", "bloqueio"].includes(value)) return "is-away";
  if (["H1", "H2", "H3", "4A", "4B"].includes(value)) return "is-work";
  return "is-empty";
}

function formatMonthYear(monthValue) {
  const [year, month] = monthValue.split("-").map(Number);
  return new Date(year, month - 1, 1).toLocaleDateString("pt-BR", {
    month: "long",
    year: "numeric",
  });
}

function renderPreferenceSummary(employee) {
  const entries =
    employee.workload === "6h"
      ? [
          ["H1", getShiftPreference(employee, "H1")],
          ["H2", getShiftPreference(employee, "H2")],
          ["H3", getShiftPreference(employee, "H3")],
        ]
      : [
          ["H4", getShiftPreference(employee, "4A")],
          ["H5", getShiftPreference(employee, "4B")],
        ];

  return entries.map(([label, value]) => `${label}: ${value}%`).join(" · ");
}

function renderPreferenceEditor(employee) {
  const preferences = employee.preferences || { H1: 0, H2: 0, H3: 0, "4A": 0, "4B": 0 };
  const fields =
    employee.workload === "6h"
      ? [
          ["H1", "H1"],
          ["H2", "H2"],
          ["H3", "H3"],
        ]
      : [
          ["4A", "H4"],
          ["4B", "H5"],
        ];

  return fields
    .map(
      ([key, label]) => `
        <label class="pref-mini">
          <span>${label}</span>
          <input data-id="${employee.id}" data-pref="${key}" type="number" min="0" max="100" value="${Number(preferences[key] || 0)}" />
        </label>
      `
    )
    .join("");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

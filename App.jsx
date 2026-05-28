import { useState, useEffect, useMemo } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";

const DEFAULT_PASS = "fundacion2024";
const PAGE_SIZE = 30;
// Chars sin ambigüedad (sin O, 0, I, 1)
const CODE_CHARS = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";

// ── Helpers puros ─────────────────────────────────────────────────────────────
function genId() {
  return Date.now().toString(36) + Math.random().toString(36).substring(2, 8);
}
function genUniqueCode(usedCodes) {
  let code, attempts = 0;
  do {
    code = Array.from({ length: 6 }, () => CODE_CHARS[Math.floor(Math.random() * CODE_CHARS.length)]).join("");
    if (++attempts > 10000) throw new Error("No se pudo generar código único");
  } while (usedCodes.has(code));
  usedCodes.add(code);
  return code;
}
function isBlank(val) {
  return val === null || val === undefined || String(val).trim() === "" || val === "undefined";
}
function getWorkerName(w) {
  const nf = ["nombre","Nombre","NOMBRE","name","Name"];
  const af = ["apellidos","Apellidos","APELLIDOS","apellido","Apellido","surname"];
  const n = nf.map(f => w.data[f]).find(v => v && String(v).trim()) || "";
  const a = af.map(f => w.data[f]).find(v => v && String(v).trim()) || "";
  return [n, a].filter(Boolean).join(" ") || w.code;
}
// Cada trabajador tiene sus propias columnas (las de su colegio)
function getWorkerColumns(w, sources) {
  const src = sources.find(s => s.id === w.sourceId);
  return src ? src.columns : Object.keys(w.data);
}
function getMissing(w, sources, submissions) {
  const cols = getWorkerColumns(w, sources);
  const sub = submissions[w.code];
  const d = sub ? sub.data : w.data;
  return cols.filter(c => isBlank(d[c])).length;
}
function mergeAllColumns(sources) {
  const seen = new Set();
  const result = [];
  for (const s of sources) {
    for (const col of (s.columns || [])) {
      if (!seen.has(col)) { seen.add(col); result.push(col); }
    }
  }
  return result;
}
function guessColumnType(col) {
  const n = col.toLowerCase();
  if (/email|correo|mail/.test(n)) return { icon: "✉️", label: "Email" };
  if (/tel[eé]|m[oó]vil|phone|celular/.test(n)) return { icon: "📱", label: "Teléfono" };
  if (/^dni$|^nif$|documento|identidad/.test(n)) return { icon: "🪪", label: "Documento" };
  if (/nombre|name/.test(n) && !/apellido/.test(n)) return { icon: "👤", label: "Nombre" };
  if (/apellido|surname/.test(n)) return { icon: "👤", label: "Apellido" };
  if (/fecha|date|nacimiento|alta|baja/.test(n)) return { icon: "📅", label: "Fecha" };
  if (/cargo|puesto|rol|job/.test(n)) return { icon: "💼", label: "Cargo" };
  if (/departamento|dept|area|área|centro|colegio|escuela/.test(n)) return { icon: "🏢", label: "Centro" };
  if (/direcci[oó]n|address|domicilio/.test(n)) return { icon: "📍", label: "Dirección" };
  if (/ciudad|localidad|municipio|city/.test(n)) return { icon: "🏙️", label: "Ciudad" };
  if (/c[oó]digo postal|cp|zip/.test(n)) return { icon: "🔢", label: "C.P." };
  if (/banco|iban|cuenta|bank/.test(n)) return { icon: "🏦", label: "Banco" };
  return { icon: "📝", label: "Campo" };
}

// ── Storage con localStorage ─────────────────────────────────────────────────
const _mem = new Map();

async function storageSave(key, value) {
  _mem.set(key, value);
  try { localStorage.setItem(key, value); } catch {}
  return true;
}

async function storageLoad(key) {
  try {
    const v = localStorage.getItem(key);
    if (v !== null) { _mem.set(key, v); return v; }
  } catch {}
  return _mem.get(key) || null;
}

async function saveAllData(workers, sources, onProgress) {
  if (onProgress) onProgress(1, 1);
  await storageSave("portal-workers", JSON.stringify({ workers, sources }));
}

async function loadAllData() {
  const v = await storageLoad("portal-workers");
  if (v) {
    try {
      const p = JSON.parse(v);
      return { workers: p.workers || [], sources: p.sources || [] };
    } catch {}
  }
  // Migrar datos del formato antiguo (chunks)
  const metaRaw = await storageLoad("wdata-meta");
  if (metaRaw) {
    try {
      const { chunks = 0, sources = [] } = JSON.parse(metaRaw);
      let workers = [];
      for (let i = 0; i < chunks; i++) {
        const c = await storageLoad(`wdata-chunk-${i}`);
        if (c) workers = workers.concat(JSON.parse(c));
      }
      if (workers.length) {
        await storageSave("portal-workers", JSON.stringify({ workers, sources }));
        return { workers, sources };
      }
    } catch {}
  }
  return { workers: [], sources: [] };
}

async function saveSubmission(code, data) {
  await storageSave(`sub-${code}`, JSON.stringify({ data, ts: new Date().toISOString() }));
}

async function loadSubmissions(workers) {
  const subs = {};
  await Promise.all(workers.map(async w => {
    const r = await storageLoad(`sub-${w.code}`);
    if (r) { try { subs[w.code] = JSON.parse(r); } catch {} }
  }));
  return subs;
}

// ── Parser de archivo ─────────────────────────────────────────────────────────
function parseFile(file) {
  return new Promise((resolve, reject) => {
    const ext = file.name.split(".").pop().toLowerCase();
    const reader = new FileReader();
    if (ext === "csv") {
      reader.onload = e => Papa.parse(e.target.result, {
        header: true, skipEmptyLines: true,
        complete: r => resolve({ columns: r.meta.fields || [], rows: r.data }),
        error: reject,
      });
      reader.readAsText(file, "UTF-8");
    } else {
      reader.onload = e => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array", cellDates: true });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
          const columns = rows.length > 0 ? Object.keys(rows[0]) : [];
          const cleanRows = rows.map(r => {
            const out = {};
            for (const k of columns) {
              const v = r[k];
              out[k] = v instanceof Date ? v.toLocaleDateString("es-ES") : String(v ?? "").trim();
            }
            return out;
          });
          resolve({ columns, rows: cleanRows });
        } catch (err) { reject(err); }
      };
      reader.readAsArrayBuffer(file);
    }
  });
}

// ── Descargas ─────────────────────────────────────────────────────────────────
function downloadExcel(rows, sheetName, filename) {
  if (!rows.length) { alert("No hay datos para exportar."); return; }
  const ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = Object.keys(rows[0]).map(k => ({
    wch: Math.max(k.length, ...rows.slice(0, 50).map(r => String(r[k] ?? "").length), 10)
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  const b64 = XLSX.write(wb, { bookType: "xlsx", type: "base64" });
  const a = Object.assign(document.createElement("a"), {
    href: "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + b64,
    download: filename,
  });
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
}
function downloadCsv(rows, fields, filename) {
  if (!rows.length) { alert("No hay datos para exportar."); return; }
  const csv = fields ? Papa.unparse({ fields, data: rows }) : Papa.unparse(rows);
  const a = Object.assign(document.createElement("a"), {
    href: "data:text/csv;charset=utf-8,\uFEFF" + encodeURIComponent(csv),
    download: filename,
  });
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
}

// ── Matching de trabajadores ─────────────────────────────────────────────────
function norm(v) {
  return String(v || "").toLowerCase().trim()
    .replace(/[áàä]/g,"a").replace(/[éèë]/g,"e").replace(/[íìï]/g,"i")
    .replace(/[óòö]/g,"o").replace(/[úùü]/g,"u").replace(/ñ/g,"n");
}
// Devuelve el índice del trabajador existente que coincide con la fila, o -1
function findMatch(data, workerList, existingSourceId) {
  const dniFields  = ["DNI","dni","NIF","nif","Dni"];
  const emailFields = ["email","Email","EMAIL","correo","Correo"];
  const nameFields  = ["nombre","Nombre","NOMBRE","name","Name"];
  const surnameFields = ["apellidos","Apellidos","APELLIDOS","apellido","Apellido"];

  const dataDni   = dniFields.map(f => data[f]).find(v => v && v.trim());
  const dataEmail = emailFields.map(f => data[f]).find(v => v && v.trim());
  const dataName  = norm(nameFields.map(f => data[f]).find(v => v) || "");
  const dataSurn  = norm(surnameFields.map(f => data[f]).find(v => v) || "");

  // 1. DNI exacto (más fiable)
  if (dataDni) {
    const idx = workerList.findIndex(w => dniFields.some(f => norm(w.data[f]) === norm(dataDni)));
    if (idx !== -1) return { idx, method: "DNI" };
  }
  // 2. Email exacto
  if (dataEmail) {
    const idx = workerList.findIndex(w => emailFields.some(f => norm(w.data[f]) === norm(dataEmail)));
    if (idx !== -1) return { idx, method: "email" };
  }
  // 3. Nombre + apellidos dentro del mismo colegio
  if (dataName && dataSurn) {
    const idx = workerList.findIndex(w =>
      w.sourceId === existingSourceId &&
      norm(nameFields.map(f => w.data[f]).find(v => v) || "") === dataName &&
      norm(surnameFields.map(f => w.data[f]).find(v => v) || "") === dataSurn
    );
    if (idx !== -1) return { idx, method: "nombre+apellidos" };
  }
  return null;
}
// Analiza un conjunto de filas y devuelve estadísticas sin modificar nada
function analyzeRows(rows, includedCols, existingWorkers, existingSourceId) {
  const toUpdate = [], toAdd = [], duplicateWarnings = [];
  const matchedIdxs = new Set();
  for (const row of rows) {
    const data = {};
    includedCols.forEach(col => { data[col] = row[col] ?? ""; });
    const match = findMatch(data, existingWorkers, existingSourceId);
    if (match) {
      if (matchedIdxs.has(match.idx)) {
        duplicateWarnings.push({ data, reason: `Coincide con un trabajador ya procesado (${match.method})` });
      } else {
        matchedIdxs.add(match.idx);
        toUpdate.push({ data, method: match.method });
      }
    } else {
      toAdd.push(data);
    }
  }
  return { toUpdate, toAdd, duplicateWarnings };
}

// ── Componente principal ──────────────────────────────────────────────────────
export default function App() {
  const [screen, setScreen] = useState("home");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  // Auth - multi-cuenta
  const [adminAccounts, setAdminAccounts] = useState([]); // [{ id, name, password, isSuperAdmin }]
  const [currentAdmin, setCurrentAdmin] = useState(null); // cuenta activa
  const [adminNameInput, setAdminNameInput] = useState("");
  const [adminPassInput, setAdminPassInput] = useState("");
  const [showPass, setShowPass] = useState(false);
  const [loginAttempts, setLoginAttempts] = useState(0);
  const [lockUntil, setLockUntil] = useState(0);

  // Data
  const [workers, setWorkers] = useState([]);
  const [sources, setSources] = useState([]); // [{ id, name, columns, importedAt, count }]
  const [submissions, setSubmissions] = useState({});

  // Worker
  const [codeInput, setCodeInput] = useState("");
  const [activeWorker, setActiveWorker] = useState(null);
  const [formData, setFormData] = useState({});
  const [submitAttempted, setSubmitAttempted] = useState(false);

  // Admin
  const [adminTab, setAdminTab] = useState("workers");
  const [expandedId, setExpandedId] = useState(null);
  const [selectedWorkers, setSelectedWorkers] = useState(new Set()); // ids seleccionados para borrado masivo
  const [copiedId, setCopiedId] = useState("");
  const [msgModal, setMsgModal] = useState(null);
  const [newPassInput, setNewPassInput] = useState("");
  const [showNewPass, setShowNewPass] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [sourceFilter, setSourceFilter] = useState("all");
  const [currentPage, setCurrentPage] = useState(1);
  // Gestión de cuentas admin
  const [newAccountName, setNewAccountName] = useState("");
  const [newAccountPass, setNewAccountPass] = useState("");
  const [showNewAccountPass, setShowNewAccountPass] = useState(false);

  // Import
  const [importStep, setImportStep] = useState("name"); // name | preview | saving
  const [importAnalysis, setImportAnalysis] = useState(null); // { toUpdate, toAdd, duplicateWarnings }
  const [importSourceName, setImportSourceName] = useState("");
  const [previewData, setPreviewData] = useState(null);
  const [columnConfig, setColumnConfig] = useState({});
  const [importProgress, setImportProgress] = useState("");
  const [importError, setImportError] = useState("");

  // Re-analizar cuando cambia la selección de columnas
  useEffect(() => {
    if (!previewData) return;
    const includedCols = previewData.columns.filter(c => columnConfig[c] !== false);
    const existingSource = sources.find(s => s.name === importSourceName.trim());
    const analysis = analyzeRows(previewData.rows, includedCols, workers, existingSource?.id || null);
    setImportAnalysis(analysis);
  }, [columnConfig, previewData, importSourceName]);

  // ── Carga inicial ───────────────────────────────────────────────────────────
  useEffect(() => {
    const safetyTimer = setTimeout(() => setLoading(false), 5000);

    (async () => {
      try {
        // Cuentas de admin
        const defaultAccounts = [{ id: genId(), name: "Administrador", password: DEFAULT_PASS, isSuperAdmin: true }];
        try {
          const acc = await window.storage.get("admin-accounts");
          setAdminAccounts(acc ? JSON.parse(acc.value) : defaultAccounts);
          if (!acc) window.storage.set("admin-accounts", JSON.stringify(defaultAccounts)).catch(() => {});
        } catch { setAdminAccounts(defaultAccounts); }

        // Trabajadores
        const { workers: ws, sources: srcs } = await loadAllData();
        if (ws.length) {
          setWorkers(ws);
          setSources(srcs);
          setSubmissions(await loadSubmissions(ws));
        }
      } catch {}

      clearTimeout(safetyTimer);
      setLoading(false);
    })();

    return () => clearTimeout(safetyTimer);
  }, []);

  // ── Stats ───────────────────────────────────────────────────────────────────
  const stats = useMemo(() => {
    const total = workers.length;
    const responded = Object.keys(submissions).length;
    const complete = workers.filter(w => getMissing(w, sources, submissions) === 0).length;
    return { total, responded, complete, pending: total - responded, pct: total ? Math.round(complete / total * 100) : 0 };
  }, [workers, sources, submissions]);

  // ── Trabajadores filtrados y paginados ─────────────────────────────────────
  const filteredWorkers = useMemo(() => {
    return workers.filter(w => {
      const hasSub = !!submissions[w.code];
      const missing = getMissing(w, sources, submissions);
      if (sourceFilter !== "all" && w.sourceId !== sourceFilter) return false;
      if (statusFilter === "pending" && hasSub) return false;
      if (statusFilter === "submitted" && !hasSub) return false;
      if (statusFilter === "complete" && missing !== 0) return false;
      if (!searchQuery) return true;
      const q = searchQuery.toLowerCase();
      return (
        getWorkerName(w).toLowerCase().includes(q) ||
        w.code.toLowerCase().includes(q) ||
        Object.values(w.data).some(v => v && String(v).toLowerCase().includes(q))
      );
    });
  }, [workers, sources, submissions, statusFilter, sourceFilter, searchQuery]);

  const totalPages = Math.max(1, Math.ceil(filteredWorkers.length / PAGE_SIZE));
  const pagedWorkers = useMemo(() => {
    const start = (currentPage - 1) * PAGE_SIZE;
    return filteredWorkers.slice(start, start + PAGE_SIZE);
  }, [filteredWorkers, currentPage]);

  useEffect(() => { setCurrentPage(1); }, [searchQuery, statusFilter, sourceFilter]);

  // ── Login admin ─────────────────────────────────────────────────────────────
  const handleAdminLogin = () => {
    if (Date.now() < lockUntil) { setError(`Demasiados intentos. Espera ${Math.ceil((lockUntil - Date.now()) / 1000)}s.`); return; }
    const account = adminAccounts.find(a => a.name.toLowerCase() === adminNameInput.trim().toLowerCase() && a.password === adminPassInput);
    if (account) {
      setError(""); setAdminPassInput(""); setAdminNameInput(""); setLoginAttempts(0);
      setCurrentAdmin(account); setScreen("admin");
    } else {
      const attempts = loginAttempts + 1;
      setLoginAttempts(attempts);
      if (attempts >= 5) { setLockUntil(Date.now() + 30000); setError("5 intentos fallidos. Bloqueado 30 segundos."); }
      else setError(`Usuario o contraseña incorrectos (${attempts}/5)`);
    }
  };

  // ── Login trabajador ────────────────────────────────────────────────────────
  const [workerAttempts, setWorkerAttempts] = useState(0);
  const [workerLock, setWorkerLock] = useState(0);
  const handleWorkerLogin = () => {
    if (Date.now() < workerLock) { setError(`Demasiados intentos. Espera ${Math.ceil((workerLock - Date.now()) / 1000)}s.`); return; }
    const code = codeInput.toUpperCase().trim();
    const w = workers.find(w => w.code === code);
    if (!w) {
      const att = workerAttempts + 1;
      setWorkerAttempts(att);
      if (att >= 10) { setWorkerLock(Date.now() + 60000); setError("Demasiados intentos fallidos. Espera 60 segundos."); }
      else setError("Código no encontrado. Comprueba el código e inténtalo de nuevo.");
      return;
    }
    setError(""); setWorkerAttempts(0);
    setActiveWorker(w);
    const sub = submissions[w.code];
    // FIX: mostrar datos de la submission si existe, si no los originales
    setFormData(sub ? { ...sub.data } : { ...w.data });
    setSubmitAttempted(false);
    setScreen("worker");
  };

  // ── Selección de archivo ────────────────────────────────────────────────────
  const handleFileSelected = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setImportStep("preview");
    setImportProgress("Leyendo archivo...");
    setImportError("");
    try {
      const { columns, rows } = await parseFile(file);
      const cfg = {};
      columns.forEach(col => { cfg[col] = true; });
      setColumnConfig(cfg);
      setPreviewData({ columns, rows, fileName: file.name });
      setImportProgress("");
    } catch (err) {
      alert("Error al leer el archivo: " + err.message);
      setImportStep("name");
    }
    e.target.value = "";
  };

  // ── Confirmar importación ───────────────────────────────────────────────────
  const handleConfirmImport = async () => {
    if (!previewData || !importSourceName.trim()) return;
    setImportStep("saving");
    setImportProgress("Procesando filas...");

    try {
      const sourceName = importSourceName.trim();
      const includedCols = previewData.columns.filter(c => columnConfig[c] !== false);
      const existingSource = sources.find(s => s.name === sourceName);
      const usedCodes = new Set(workers.map(w => w.code));

      let updatedWorkers = [...workers];
      let newCount = 0, updatedCount = 0;
      const matchedIdxs = new Set();

      for (const row of previewData.rows) {
        const data = {};
        includedCols.forEach(col => { data[col] = row[col] ?? ""; });
        const match = findMatch(data, updatedWorkers, existingSource?.id || null);
        const existingIdx = (match && !matchedIdxs.has(match.idx)) ? match.idx : -1;
        if (existingIdx !== -1) matchedIdxs.add(existingIdx);
        if (existingIdx !== -1) {
          updatedWorkers[existingIdx] = {
            ...updatedWorkers[existingIdx],
            data: { ...updatedWorkers[existingIdx].data, ...data },
            sourceId: existingSource?.id || updatedWorkers[existingIdx].sourceId,
          };
          updatedCount++;
        } else {
          updatedWorkers.push({ id: genId(), code: genUniqueCode(usedCodes), sourceId: existingSource?.id || null, data });
          newCount++;
        }
      }

      let updatedSources;
      if (existingSource) {
        updatedSources = sources.map(s => s.id === existingSource.id
          ? { ...s, columns: includedCols, importedAt: new Date().toLocaleDateString("es-ES") }
          : s
        );
      } else {
        const newSourceId = genId();
        updatedSources = [...sources, { id: newSourceId, name: sourceName, columns: includedCols, importedAt: new Date().toLocaleDateString("es-ES"), count: newCount }];
        updatedWorkers = updatedWorkers.map(w => w.sourceId === null ? { ...w, sourceId: newSourceId } : w);
      }

      setImportProgress("Guardando datos...");
      await saveAllData(updatedWorkers, updatedSources, () => {
        setImportProgress("Guardando datos... ✓");
      });

      setWorkers(updatedWorkers);
      setSources(updatedSources);
      setPreviewData(null);
      setImportSourceName("");
      setImportAnalysis(null);
      setAdminTab("workers");
      alert(`✅ Importación completada\n• ${newCount} trabajadores nuevos\n• ${updatedCount} actualizados`);

    } catch (err) {
      console.error("Import error:", err);
      setImportError(`${err?.message || String(err)}`);
      setImportStep("preview"); // volver a preview para que vea el error
    } finally {
      setImportStep("name");
      setImportProgress("");
    }
  };


  // ── Envío trabajador ────────────────────────────────────────────────────────
  const handleWorkerSubmit = async () => {
    setSubmitAttempted(true);
    // FIX: validar que los campos vacíos del trabajador estén rellenados
    const workerCols = getWorkerColumns(activeWorker, sources);
    const stillEmpty = workerCols.filter(c => isBlank(formData[c]));
    if (stillEmpty.length > 0) {
      // Scroll al primer campo vacío
      document.getElementById(`field-${stillEmpty[0]}`)?.scrollIntoView({ behavior: "smooth", block: "center" });
      return;
    }
    await saveSubmission(activeWorker.code, formData);
    setSubmissions(prev => ({ ...prev, [activeWorker.code]: { data: formData, ts: new Date().toISOString() } }));
    setScreen("done");
  };

  // ── Exports ─────────────────────────────────────────────────────────────────
  const buildDataRows = () => {
    const allCols = mergeAllColumns(sources);
    return workers.map(w => {
      const sub = submissions[w.code];
      const d = sub ? sub.data : w.data;
      const src = sources.find(s => s.id === w.sourceId);
      const row = { Colegio: src?.name || "" };
      allCols.forEach(col => { row[col] = d[col] ?? ""; });
      return row;
    });
  };
  const buildCodeRows = () => {
    const emailFields = ["email","Email","EMAIL","correo","Correo"];
    return workers.map(w => {
      const src = sources.find(s => s.id === w.sourceId);
      return {
        Nombre: getWorkerName(w),
        Email: emailFields.map(f => w.data[f]).find(Boolean) || "",
        Código: w.code,
        Colegio: src?.name || "",
        Estado: getMissing(w, sources, submissions) === 0 ? "Completo" : submissions[w.code] ? "Enviado (incompleto)" : "Pendiente",
        "Campos vacíos": getMissing(w, sources, submissions),
      };
    });
  };

  const handleExportDataXlsx  = () => downloadExcel(buildDataRows(), "Trabajadores", "trabajadores_actualizado.xlsx");
  const handleExportDataCsv   = () => downloadCsv(buildDataRows(), null, "trabajadores_actualizado.csv");
  const handleExportCodesXlsx = () => downloadExcel(buildCodeRows(), "Códigos", "codigos_trabajadores.xlsx");
  const handleExportCodesCsv  = () => downloadCsv(buildCodeRows(), null, "codigos_trabajadores.csv");

  // ── Copiar mensaje ──────────────────────────────────────────────────────────
  const copyCode = (w) => {
    const name = getWorkerName(w).split(" ")[0] || "";
    const msg = `Hola${name ? " " + name : ""},\n\nPara actualizar tus datos en el nuevo sistema, accede al portal e introduce tu código personal:\n\n🔑 Código: ${w.code}\n\nGracias.`;
    try {
      navigator.clipboard.writeText(msg)
        .then(() => { setCopiedId(w.id); setTimeout(() => setCopiedId(""), 2000); })
        .catch(() => setMsgModal({ name: getWorkerName(w), msg }));
    } catch { setMsgModal({ name: getWorkerName(w), msg }); }
  };


  const deleteWorker = async (w) => {
    const name = getWorkerName(w);
    if (!window.confirm(`¿Eliminar a "${name}"?\n\nEsta acción no se puede deshacer.`)) return;
    const updated = workers.filter(x => x.id !== w.id);
    const updatedSources = sources.map(s =>
      s.id === w.sourceId ? { ...s, count: updated.filter(x => x.sourceId === s.id).length } : s
    );
    await saveAllData(updated, updatedSources);
    try { await storageSave(`sub-${w.code}`, ""); } catch {}
    setWorkers(updated);
    setSources(updatedSources);
    const s2 = { ...submissions }; delete s2[w.code]; setSubmissions(s2);
    if (expandedId === w.id) setExpandedId(null);
  };

  const deleteSource = async (sourceId) => {
    const src = sources.find(s => s.id === sourceId);
    const count = workers.filter(w => w.sourceId === sourceId).length;
    if (!window.confirm(`¿Eliminar "${src?.name}" y sus ${count} trabajadores?\n\nEsta acción no se puede deshacer.`)) return;
    const updated = workers.filter(w => w.sourceId !== sourceId);
    const updatedSources = sources.filter(s => s.id !== sourceId);
    await saveAllData(updated, updatedSources);
    setWorkers(updated); setSources(updatedSources);
    const s2 = { ...submissions };
    workers.filter(w => w.sourceId === sourceId).forEach(w => { delete s2[w.code]; });
    setSubmissions(s2);
    if (sourceFilter === sourceId) setSourceFilter("all");
  };

  const toggleSelectWorker = (id) => {
    setSelectedWorkers(prev => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  };

  const toggleSelectAll = () => {
    if (selectedWorkers.size === pagedWorkers.length) {
      setSelectedWorkers(new Set());
    } else {
      setSelectedWorkers(new Set(pagedWorkers.map(w => w.id)));
    }
  };

  const deleteSelected = async () => {
    if (!selectedWorkers.size) return;
    if (!window.confirm(`¿Eliminar ${selectedWorkers.size} trabajador${selectedWorkers.size !== 1 ? "es" : ""}?\n\nEsta acción no se puede deshacer.`)) return;
    const updated = workers.filter(w => !selectedWorkers.has(w.id));
    const updatedSources = sources.map(s => ({ ...s, count: updated.filter(x => x.sourceId === s.id).length }));
    await saveAllData(updated, updatedSources);
    setWorkers(updated);
    setSources(updatedSources);
    const s2 = { ...submissions };
    workers.filter(w => selectedWorkers.has(w.id)).forEach(w => { delete s2[w.code]; });
    setSubmissions(s2);
    setSelectedWorkers(new Set());
  };

  const changeAdminPass = async () => {
    if (!newPassInput.trim() || !currentAdmin) return;
    const updated = adminAccounts.map(a => a.id === currentAdmin.id ? { ...a, password: newPassInput.trim() } : a);
    await window.storage.set("admin-accounts", JSON.stringify(updated));
    setAdminAccounts(updated);
    setCurrentAdmin(prev => ({ ...prev, password: newPassInput.trim() }));
    setNewPassInput("");
    alert("Contraseña actualizada correctamente.");
  };

  const createAdminAccount = async () => {
    if (!newAccountName.trim() || !newAccountPass.trim()) return;
    if (adminAccounts.some(a => a.name.toLowerCase() === newAccountName.trim().toLowerCase())) {
      alert("Ya existe una cuenta con ese nombre."); return;
    }
    const newAccount = { id: genId(), name: newAccountName.trim(), password: newAccountPass.trim(), isSuperAdmin: false };
    const updated = [...adminAccounts, newAccount];
    await window.storage.set("admin-accounts", JSON.stringify(updated));
    setAdminAccounts(updated);
    setNewAccountName(""); setNewAccountPass("");
    alert(`Cuenta "${newAccount.name}" creada correctamente.`);
  };

  const deleteAdminAccount = async (id) => {
    if (id === currentAdmin?.id) { alert("No puedes eliminar tu propia cuenta."); return; }
    const target = adminAccounts.find(a => a.id === id);
    if (!window.confirm(`¿Eliminar la cuenta de "${target?.name}"?`)) return;
    const updated = adminAccounts.filter(a => a.id !== id);
    await window.storage.set("admin-accounts", JSON.stringify(updated));
    setAdminAccounts(updated);
  };

  // ── RENDER ────────────────────────────────────────────────────────────────
  if (loading) return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center">
      <div className="flex flex-col items-center gap-3 text-slate-400">
        <div className="w-6 h-6 border-2 border-slate-300 border-t-indigo-500 rounded-full animate-spin" />
        <span className="text-sm">Cargando sistema...</span>
      </div>
    </div>
  );

  // ── HOME ──────────────────────────────────────────────────────────────────
  if (screen === "home") return (
    <div className="min-h-screen bg-gradient-to-br from-indigo-50 via-white to-violet-50 flex items-center justify-center p-4">
      <div className="w-full max-w-md">
        <div className="text-center mb-8">
          <div className="inline-flex items-center justify-center w-16 h-16 bg-indigo-600 rounded-2xl shadow-lg mb-4">
            <svg className="w-8 h-8 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" /></svg>
          </div>
          <h1 className="text-2xl font-bold text-slate-900">Portal de Datos</h1>
          <p className="text-slate-500 text-sm mt-1">Actualización de expedientes del personal</p>
        </div>
        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
          <button onClick={() => { setError(""); setScreen("admin-login"); }} className="w-full flex items-center gap-4 p-5 hover:bg-slate-50 transition-colors border-b border-slate-100">
            <div className="w-10 h-10 bg-indigo-100 rounded-xl flex items-center justify-center shrink-0">
              <svg className="w-5 h-5 text-indigo-600" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" /></svg>
            </div>
            <div className="text-left flex-1"><div className="font-semibold text-slate-900 text-sm">Acceso Administrador</div><div className="text-slate-400 text-xs">Gestión de datos y códigos</div></div>
            <svg className="w-4 h-4 text-slate-300" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" /></svg>
          </button>
          <button onClick={() => { setError(""); setScreen("worker-login"); }} className="w-full flex items-center gap-4 p-5 hover:bg-slate-50 transition-colors">
            <div className="w-10 h-10 bg-violet-100 rounded-xl flex items-center justify-center shrink-0">
              <svg className="w-5 h-5 text-violet-600" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" /></svg>
            </div>
            <div className="text-left flex-1"><div className="font-semibold text-slate-900 text-sm">Soy trabajador</div><div className="text-slate-400 text-xs">Tengo un código personal</div></div>
            <svg className="w-4 h-4 text-slate-300" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" /></svg>
          </button>
        </div>
      </div>
    </div>
  );

  // ── ADMIN LOGIN ────────────────────────────────────────────────────────────
  if (screen === "admin-login") return (
    <div className="min-h-screen bg-gradient-to-br from-indigo-50 via-white to-violet-50 flex items-center justify-center p-4">
      <div className="w-full max-w-sm">
        <button onClick={() => setScreen("home")} className="flex items-center gap-1 text-slate-400 hover:text-slate-600 text-sm mb-6 transition-colors">
          <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" /></svg> Volver
        </button>
        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-7">
          <h2 className="text-xl font-bold text-slate-900 mb-1">Administrador</h2>
          <p className="text-slate-400 text-sm mb-6">Introduce tu nombre de usuario y contraseña.</p>
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1.5">Nombre de usuario</label>
              <input type="text" value={adminNameInput}
                onChange={e => { setAdminNameInput(e.target.value); setError(""); }}
                onKeyDown={e => e.key === "Enter" && document.getElementById("admin-pass-input")?.focus()}
                disabled={Date.now() < lockUntil}
                className="w-full border border-slate-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 disabled:opacity-50"
                placeholder="Ej: Administrador" autoFocus />
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1.5">Contraseña</label>
              <div className="relative">
                <input id="admin-pass-input" type={showPass ? "text" : "password"} value={adminPassInput}
                  onChange={e => { setAdminPassInput(e.target.value); setError(""); }}
                  onKeyDown={e => e.key === "Enter" && handleAdminLogin()}
                  disabled={Date.now() < lockUntil}
                  className="w-full border border-slate-200 rounded-xl px-4 py-2.5 pr-11 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 disabled:opacity-50" placeholder="••••••••••" />
                <button type="button" onClick={() => setShowPass(v => !v)} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600">
                  {showPass
                    ? <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858.908a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242M9.88 9.88l-3.29-3.29m7.532 7.532l3.29 3.29M3 3l3.59 3.59m0 0A9.953 9.953 0 0112 5c4.478 0 8.268 2.943 9.543 7a10.025 10.025 0 01-4.132 5.411m0 0L21 21" /></svg>
                    : <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" /><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" /></svg>
                  }
                </button>
              </div>
              {adminAccounts.length === 0 && <p className="text-xs text-indigo-500 font-medium mt-1.5">Usuario: <span className="font-mono">Administrador</span> · Contraseña: <span className="font-mono">fundacion2024</span></p>}
            </div>
            {error && <div className="bg-red-50 text-red-600 text-sm px-3 py-2 rounded-lg">{error}</div>}
            <button onClick={handleAdminLogin} disabled={Date.now() < lockUntil || !adminNameInput.trim() || !adminPassInput} className="w-full bg-indigo-600 text-white py-2.5 rounded-xl font-semibold text-sm hover:bg-indigo-700 disabled:opacity-50 transition-colors">Entrar al panel</button>
          </div>
        </div>
      </div>
    </div>
  );

  // ── WORKER LOGIN ───────────────────────────────────────────────────────────
  if (screen === "worker-login") return (
    <div className="min-h-screen bg-gradient-to-br from-indigo-50 via-white to-violet-50 flex items-center justify-center p-4">
      <div className="w-full max-w-sm">
        <button onClick={() => setScreen("home")} className="flex items-center gap-1 text-slate-400 hover:text-slate-600 text-sm mb-6 transition-colors">
          <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" /></svg> Volver
        </button>
        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-7">
          <h2 className="text-xl font-bold text-slate-900 mb-1">Introduce tu código</h2>
          <p className="text-slate-400 text-sm mb-6">Tu responsable te habrá facilitado un código personal de 6 caracteres.</p>
          <div className="space-y-4">
            <input type="text" value={codeInput}
              onChange={e => { setCodeInput(e.target.value.toUpperCase().replace(/[^A-Z0-9]/g, "")); setError(""); }}
              onKeyDown={e => e.key === "Enter" && codeInput.length === 6 && handleWorkerLogin()}
              disabled={Date.now() < workerLock}
              className="w-full border-2 border-slate-200 rounded-xl px-4 py-4 text-center text-3xl font-mono font-bold tracking-[0.3em] focus:outline-none focus:border-indigo-500 transition-colors disabled:opacity-50"
              maxLength={6} placeholder="XXXXXX" autoFocus />
            {error && <div className="bg-red-50 text-red-600 text-sm px-3 py-2 rounded-lg">{error}</div>}
            <button onClick={handleWorkerLogin} disabled={codeInput.length < 6 || Date.now() < workerLock}
              className="w-full bg-violet-600 text-white py-3 rounded-xl font-semibold text-sm hover:bg-violet-700 disabled:opacity-40 disabled:cursor-not-allowed transition-colors">
              Ver mis datos →
            </button>
          </div>
        </div>
      </div>
    </div>
  );

  // ── WORKER FORM ────────────────────────────────────────────────────────────
  if (screen === "worker" && activeWorker) {
    const workerCols = getWorkerColumns(activeWorker, sources);
    const blankCols = workerCols.filter(c => isBlank(activeWorker.data[c]));
    const filledCols = workerCols.filter(c => !isBlank(activeWorker.data[c]));
    const alreadySubmitted = !!submissions[activeWorker.code];
    // FIX: validar campos al intentar enviar
    const invalidCols = submitAttempted ? workerCols.filter(c => isBlank(formData[c])) : [];

    return (
      <div className="min-h-screen bg-slate-50">
        <div className="bg-white border-b border-slate-100 sticky top-0 z-10">
          <div className="max-w-2xl mx-auto px-4 py-4 flex items-center justify-between">
            <div>
              <h1 className="font-bold text-slate-900">{getWorkerName(activeWorker)}</h1>
              <p className="text-xs text-slate-400 mt-0.5">Código: <span className="font-mono">{activeWorker.code}</span></p>
            </div>
            {alreadySubmitted && <span className="inline-flex items-center gap-1 bg-emerald-50 text-emerald-700 text-xs font-semibold px-3 py-1.5 rounded-full border border-emerald-200">✓ Ya enviado</span>}
          </div>
        </div>
        <div className="max-w-2xl mx-auto p-4 pb-8 space-y-5">
          {blankCols.length > 0 ? (
            <div className="bg-amber-50 border border-amber-200 rounded-xl p-4 flex gap-3">
              <svg className="w-5 h-5 text-amber-500 shrink-0 mt-0.5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" /></svg>
              <p className="text-amber-800 text-sm"><strong>Tienes {blankCols.length} campo{blankCols.length !== 1 ? "s" : ""} sin rellenar.</strong> Por favor, completa los campos en naranja antes de enviar.</p>
            </div>
          ) : (
            <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-4 flex gap-3">
              <svg className="w-5 h-5 text-emerald-500 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
              <p className="text-emerald-800 text-sm"><strong>Todos tus campos están completos.</strong> Confirma que los datos son correctos y envía.</p>
            </div>
          )}

          {blankCols.length > 0 && (
            <div className="bg-white rounded-2xl border-2 border-amber-300 overflow-hidden">
              <div className="bg-amber-50 px-5 py-3 border-b border-amber-200 flex items-center gap-2">
                <span>✏️</span>
                <h3 className="font-semibold text-amber-800 text-sm">Campos a completar</h3>
              </div>
              <div className="p-5 space-y-4">
                {blankCols.map(col => {
                  const isInvalid = invalidCols.includes(col);
                  return (
                    <div key={col} id={`field-${col}`}>
                      <label className="block text-sm font-medium text-slate-700 mb-1.5">
                        {col} <span className="text-red-400">*</span>
                      </label>
                      <input type="text" value={formData[col] || ""}
                        onChange={e => setFormData({ ...formData, [col]: e.target.value })}
                        className={`w-full border-2 rounded-xl px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:border-transparent bg-amber-50 ${isInvalid ? "border-red-400 focus:ring-red-400" : "border-amber-200 focus:ring-amber-400"}`}
                        placeholder={`Introduce ${col}...`} />
                      {isInvalid && <p className="text-xs text-red-500 mt-1">Este campo es obligatorio.</p>}
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* FIX: mostrar formData (datos actuales del form) en vez de activeWorker.data */}
          {filledCols.length > 0 && (
            <div className="bg-white rounded-2xl border border-slate-200 overflow-hidden">
              <div className="bg-slate-50 px-5 py-3 border-b border-slate-100 flex items-center gap-2">
                <span className="text-slate-400 text-sm">✅</span>
                <h3 className="font-semibold text-slate-500 text-sm">Datos existentes</h3>
              </div>
              <div className="divide-y divide-slate-50">
                {filledCols.map(col => (
                  <div key={col} className="px-5 py-3 flex justify-between items-center gap-4 text-sm">
                    <span className="text-slate-400 font-medium shrink-0">{col}</span>
                    <span className="text-slate-700 text-right">{formData[col] ?? activeWorker.data[col]}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          <button onClick={handleWorkerSubmit}
            className="w-full bg-indigo-600 text-white py-3.5 rounded-xl font-semibold text-sm hover:bg-indigo-700 transition-colors shadow-sm flex items-center justify-center gap-2">
            <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M5 13l4 4L19 7" /></svg>
            {alreadySubmitted ? "Actualizar mis datos" : "Confirmar y enviar datos"}
          </button>
          <p className="text-xs text-slate-400 text-center">🔒 Solo verás tus propios datos. Tu información es privada.</p>
        </div>
      </div>
    );
  }

  // ── DONE ──────────────────────────────────────────────────────────────────
  if (screen === "done") return (
    <div className="min-h-screen bg-gradient-to-br from-emerald-50 via-white to-teal-50 flex items-center justify-center p-4">
      <div className="w-full max-w-sm text-center">
        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-8">
          <div className="w-16 h-16 bg-emerald-100 rounded-full flex items-center justify-center mx-auto mb-5">
            <svg className="w-8 h-8 text-emerald-600" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
          </div>
          <h2 className="text-xl font-bold text-slate-900 mb-2">¡Datos enviados!</h2>
          <p className="text-slate-500 text-sm mb-6">Tus datos han sido recibidos. ¡Muchas gracias!</p>
          <button onClick={() => { setScreen("home"); setCodeInput(""); setActiveWorker(null); }}
            className="w-full bg-slate-100 text-slate-700 py-2.5 rounded-xl font-medium text-sm hover:bg-slate-200 transition-colors">Volver al inicio</button>
        </div>
      </div>
    </div>
  );

  // ── ADMIN ─────────────────────────────────────────────────────────────────
  if (screen === "admin") return (
    <div className="min-h-screen bg-slate-50">
      <div className="bg-indigo-700 text-white">
        <div className="max-w-5xl mx-auto px-4 py-4 flex items-center justify-between">
          <div>
            <h1 className="font-bold text-base">Panel de Administración</h1>
            <p className="text-indigo-300 text-xs">{sources.length} colegio{sources.length !== 1 ? "s" : ""} · {stats.total} trabajadores</p>
          </div>
          <div className="flex items-center gap-3">
            <div className="text-right">
              <p className="text-white text-xs font-semibold">{currentAdmin?.name}</p>
              {currentAdmin?.isSuperAdmin && <p className="text-indigo-300 text-xs">Superadmin</p>}
            </div>
            <button onClick={() => { setScreen("home"); setCurrentAdmin(null); }} className="text-indigo-300 hover:text-white text-xs border border-indigo-500 hover:border-white px-2.5 py-1 rounded-lg transition-colors">Salir</button>
          </div>
        </div>
      </div>

      <div className="max-w-5xl mx-auto px-4 py-4">
        {/* Stats */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 mb-4">
          {[
            { label: "Total", value: stats.total, color: "text-slate-900" },
            { label: "Pendientes", value: stats.pending, color: "text-amber-600" },
            { label: "Han respondido", value: stats.responded, color: "text-blue-600" },
            { label: "Completos", value: stats.complete, color: "text-emerald-600" },
          ].map(s => (
            <div key={s.label} className="bg-white rounded-xl border border-slate-100 p-4 text-center shadow-sm">
              <div className={`text-2xl font-bold ${s.color}`}>{s.value}</div>
              <div className="text-xs text-slate-400 mt-0.5">{s.label}</div>
            </div>
          ))}
        </div>

        {/* Progreso */}
        {stats.total > 0 && (
          <div className="bg-white rounded-xl border border-slate-100 p-4 shadow-sm mb-4">
            <div className="flex justify-between mb-2">
              <span className="text-sm font-semibold text-slate-700">Progreso global</span>
              <span className="text-sm font-bold text-indigo-600">{stats.pct}%</span>
            </div>
            <div className="h-3 bg-slate-100 rounded-full overflow-hidden">
              <div className="h-full bg-gradient-to-r from-indigo-500 to-emerald-500 rounded-full transition-all duration-500" style={{ width: `${stats.pct}%` }} />
            </div>
            <p className="text-xs text-slate-400 mt-2">{stats.complete} de {stats.total} trabajadores con todos los campos completos</p>
          </div>
        )}

        {/* Colegios */}
        {sources.length > 0 && (
          <div className="bg-white rounded-xl border border-slate-100 shadow-sm mb-4 overflow-hidden">
            <div className="px-4 py-3 border-b border-slate-100 flex items-center justify-between">
              <span className="text-sm font-semibold text-slate-700">Colegios importados</span>
              <span className="text-xs text-slate-400">{sources.length} centros</span>
            </div>
            <div className="divide-y divide-slate-50">
              {sources.map(src => {
                const srcWorkers = workers.filter(w => w.sourceId === src.id);
                const srcComplete = srcWorkers.filter(w => getMissing(w, sources, submissions) === 0).length;
                const pct = srcWorkers.length ? Math.round(srcComplete / srcWorkers.length * 100) : 0;
                return (
                  <div key={src.id} className="px-4 py-3 flex items-center gap-3">
                    <div className="w-8 h-8 bg-indigo-50 rounded-lg flex items-center justify-center shrink-0">
                      <span className="text-sm">🏫</span>
                    </div>
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center gap-2">
                        <span className="text-sm font-semibold text-slate-800 truncate">{src.name}</span>
                        <span className="text-xs text-slate-400 shrink-0">{srcWorkers.length} trabajadores</span>
                      </div>
                      <div className="flex items-center gap-2 mt-1">
                        <div className="flex-1 h-1.5 bg-slate-100 rounded-full overflow-hidden">
                          <div className="h-full bg-indigo-400 rounded-full" style={{ width: `${pct}%` }} />
                        </div>
                        <span className="text-xs text-slate-400 shrink-0">{pct}% completo</span>
                      </div>
                    </div>
                    <div className="flex items-center gap-2 shrink-0">
                      <span className="text-xs text-slate-400">Importado {src.importedAt}</span>
                      <button onClick={() => deleteSource(src.id)} className="text-xs px-2 py-1 rounded-lg border border-red-200 text-red-400 hover:bg-red-50 hover:border-red-400 transition-colors" title="Eliminar colegio">✕</button>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* Tabs */}
        <div className="flex bg-slate-100 rounded-xl p-1 mb-4 gap-1 overflow-x-auto">
          {[
            { id: "workers", label: "👥 Trabajadores" },
            { id: "upload", label: "📤 Importar" },
            { id: "export", label: "📥 Exportar" },
            { id: "settings", label: "⚙️ Ajustes" },
          ].map(t => (
            <button key={t.id} onClick={() => { setAdminTab(t.id); if (t.id !== "upload") { setImportStep("name"); setPreviewData(null); }}}
              className={`flex-1 whitespace-nowrap py-2 text-xs font-semibold rounded-lg transition-all ${(adminTab === t.id || (t.id === "upload" && adminTab === "preview")) ? "bg-white text-indigo-700 shadow-sm" : "text-slate-500 hover:text-slate-700"}`}>
              {t.label}
            </button>
          ))}
        </div>

        {/* ── WORKERS ── */}
        {adminTab === "workers" && (
          <div className="space-y-3 pb-8">
            {workers.length === 0 ? (
              <div className="bg-white rounded-2xl border border-dashed border-slate-300 p-10 text-center">
                <div className="text-4xl mb-3">📂</div>
                <p className="text-slate-500 text-sm font-medium mb-3">No hay trabajadores cargados</p>
                <button onClick={() => setAdminTab("upload")} className="bg-indigo-600 text-white text-sm px-4 py-2 rounded-xl font-medium hover:bg-indigo-700">Importar archivo →</button>
              </div>
            ) : (
              <>
                <div className="flex items-center gap-2 mb-1">
                  <input type="checkbox"
                    checked={pagedWorkers.length > 0 && pagedWorkers.every(w => selectedWorkers.has(w.id))}
                    onChange={toggleSelectAll}
                    className="w-4 h-4 rounded accent-indigo-600 cursor-pointer"
                    title="Seleccionar todos los de esta página" />
                  <span className="text-xs text-slate-400">Seleccionar página</span>
                </div>
                <div className="flex gap-2 flex-wrap">
                  <div className="relative flex-1 min-w-48">
                    <svg className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" /></svg>
                    <input type="text" value={searchQuery} onChange={e => setSearchQuery(e.target.value)} placeholder="Buscar nombre, DNI, código..."
                      className="w-full bg-white border border-slate-200 rounded-xl pl-9 pr-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500" />
                  </div>
                  <select value={sourceFilter} onChange={e => setSourceFilter(e.target.value)} className="bg-white border border-slate-200 rounded-xl px-3 py-2.5 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500">
                    <option value="all">Todos los colegios</option>
                    {sources.map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                  <select value={statusFilter} onChange={e => setStatusFilter(e.target.value)} className="bg-white border border-slate-200 rounded-xl px-3 py-2.5 text-sm text-slate-700 focus:outline-none focus:ring-2 focus:ring-indigo-500">
                    <option value="all">Todos</option>
                    <option value="pending">Pendientes</option>
                    <option value="submitted">Han enviado</option>
                    <option value="complete">Completos</option>
                  </select>
                </div>

                <div className="flex items-center justify-between px-1">
                  <p className="text-xs text-slate-400">{filteredWorkers.length} resultado{filteredWorkers.length !== 1 ? "s" : ""}{filteredWorkers.length > PAGE_SIZE ? ` · Página ${currentPage} de ${totalPages}` : ""}</p>
                  {selectedWorkers.size > 0 && (
                    <div className="flex items-center gap-2">
                      <span className="text-xs text-slate-500 font-medium">{selectedWorkers.size} seleccionado{selectedWorkers.size !== 1 ? "s" : ""}</span>
                      <button onClick={deleteSelected} className="text-xs bg-red-500 hover:bg-red-600 text-white px-3 py-1.5 rounded-lg font-semibold transition-colors flex items-center gap-1">
                        <svg className="w-3 h-3" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" /></svg>
                        Eliminar seleccionados
                      </button>
                      <button onClick={() => setSelectedWorkers(new Set())} className="text-xs text-slate-400 hover:text-slate-600 border border-slate-200 px-2.5 py-1.5 rounded-lg transition-colors">Cancelar</button>
                    </div>
                  )}
                </div>

                {pagedWorkers.map(w => {
                  const missing = getMissing(w, sources, submissions);
                  const hasSub = !!submissions[w.code];
                  const isComplete = missing === 0;
                  const isExpanded = expandedId === w.id;
                  const src = sources.find(s => s.id === w.sourceId);
                  return (
                    <div key={w.id} className="bg-white rounded-2xl border border-slate-100 shadow-sm overflow-hidden">
                      <div className="p-4 flex items-center gap-3">
                        <input type="checkbox" checked={selectedWorkers.has(w.id)}
                          onChange={() => toggleSelectWorker(w.id)}
                          onClick={e => e.stopPropagation()}
                          className="w-4 h-4 rounded accent-indigo-600 shrink-0 cursor-pointer" />
                        <div className={`w-9 h-9 rounded-xl flex items-center justify-center text-sm font-bold shrink-0 ${isComplete ? "bg-emerald-100 text-emerald-700" : hasSub ? "bg-blue-100 text-blue-700" : "bg-amber-100 text-amber-700"}`}>
                          {getWorkerName(w).charAt(0).toUpperCase() || "?"}
                        </div>
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2 flex-wrap">
                            <span className="font-semibold text-slate-900 text-sm">{getWorkerName(w)}</span>
                            {isComplete && <span className="text-xs bg-emerald-50 text-emerald-700 border border-emerald-200 px-2 py-0.5 rounded-full font-medium">Completo</span>}
                            {!isComplete && hasSub && <span className="text-xs bg-blue-50 text-blue-700 border border-blue-200 px-2 py-0.5 rounded-full font-medium">Enviado · {missing} vacío{missing !== 1 ? "s" : ""}</span>}
                            {!hasSub && missing > 0 && <span className="text-xs bg-amber-50 text-amber-700 border border-amber-200 px-2 py-0.5 rounded-full font-medium">{missing} vacío{missing !== 1 ? "s" : ""}</span>}
                          </div>
                          <div className="flex items-center gap-2 mt-1 flex-wrap">
                            <span className="font-mono text-xs text-indigo-600 bg-indigo-50 border border-indigo-100 px-2 py-0.5 rounded-md">{w.code}</span>
                            {src && <span className="text-xs text-slate-400">🏫 {src.name}</span>}
                            {hasSub && <span className="text-xs text-slate-400">{new Date(submissions[w.code].ts).toLocaleDateString("es-ES", { day: "numeric", month: "short" })}</span>}
                          </div>
                        </div>
                        <div className="flex gap-2 ml-1 shrink-0">
                          <button onClick={() => setExpandedId(isExpanded ? null : w.id)} className="text-xs text-slate-500 border border-slate-200 px-2.5 py-1.5 rounded-lg font-medium hover:text-slate-700">{isExpanded ? "Cerrar" : "Ver"}</button>
                          <button onClick={() => copyCode(w)} className={`text-xs px-3 py-1.5 rounded-lg font-semibold transition-all ${copiedId === w.id ? "bg-emerald-500 text-white" : "bg-indigo-600 text-white hover:bg-indigo-700"}`}>{copiedId === w.id ? "✓ Copiado" : "Copiar msg"}</button>
                          <button onClick={() => deleteWorker(w)} className="text-xs px-2.5 py-1.5 rounded-lg font-medium border border-red-200 text-red-500 hover:bg-red-50 hover:border-red-400 transition-colors" title="Eliminar trabajador">✕</button>
                        </div>
                      </div>
                      {isExpanded && (
                        <div className="border-t border-slate-100 bg-slate-50 p-4">
                          <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                            {getWorkerColumns(w, sources).map(col => {
                              const sub = submissions[w.code];
                              const val = sub ? sub.data[col] : w.data[col];
                              const blank = isBlank(val);
                              return (
                                <div key={col} className={`text-xs p-2.5 rounded-lg ${blank ? "bg-amber-50 border border-amber-200" : "bg-white border border-slate-100"}`}>
                                  <div className="text-slate-400 font-medium mb-0.5">{col}</div>
                                  <div className={blank ? "text-amber-400 italic" : "text-slate-800 font-medium"}>{blank ? "—vacío—" : val}</div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })}

                {totalPages > 1 && (
                  <div className="flex items-center justify-center gap-2 py-2">
                    <button onClick={() => setCurrentPage(p => Math.max(1, p - 1))} disabled={currentPage === 1}
                      className="px-3 py-1.5 text-sm border border-slate-200 rounded-lg bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-40">← Anterior</button>
                    {Array.from({ length: Math.min(5, totalPages) }, (_, i) => {
                      const page = totalPages <= 5 ? i + 1 : currentPage <= 3 ? i + 1 : currentPage >= totalPages - 2 ? totalPages - 4 + i : currentPage - 2 + i;
                      return <button key={page} onClick={() => setCurrentPage(page)}
                        className={`w-9 h-9 text-sm rounded-lg font-medium ${currentPage === page ? "bg-indigo-600 text-white" : "bg-white border border-slate-200 text-slate-600 hover:bg-slate-50"}`}>{page}</button>;
                    })}
                    <button onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))} disabled={currentPage === totalPages}
                      className="px-3 py-1.5 text-sm border border-slate-200 rounded-lg bg-white text-slate-600 hover:bg-slate-50 disabled:opacity-40">Siguiente →</button>
                  </div>
                )}
              </>
            )}
          </div>
        )}

        {/* ── IMPORT: paso 1 nombre ── */}
        {adminTab === "upload" && importStep === "name" && (
          <div className="pb-8 space-y-4">
            <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-6">
              <h3 className="font-bold text-slate-900 mb-1">Importar trabajadores</h3>
              <p className="text-sm text-slate-500 mb-5">
                Cada importación se registra por separado. Si ya existe ese colegio y vuelves a subir su Excel, <strong>solo se actualizarán sus trabajadores</strong> sin afectar al resto.
              </p>

              {sources.length > 0 && (
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 mb-5">
                  <p className="text-xs font-semibold text-slate-500 mb-2">Colegios ya cargados</p>
                  <div className="flex flex-wrap gap-2">
                    {sources.map(s => (
                      <button key={s.id} onClick={() => setImportSourceName(s.name)}
                        className="text-xs bg-white border border-slate-200 hover:border-indigo-400 hover:text-indigo-600 px-3 py-1.5 rounded-lg transition-colors">
                        🏫 {s.name}
                      </button>
                    ))}
                  </div>
                  <p className="text-xs text-slate-400 mt-2">Pulsa uno para reimportar ese colegio</p>
                </div>
              )}

              <div className="space-y-4">
                <div>
                  <label className="block text-sm font-medium text-slate-700 mb-1.5">
                    Nombre del colegio / centro <span className="text-red-400">*</span>
                  </label>
                  <input type="text" value={importSourceName} onChange={e => setImportSourceName(e.target.value)}
                    className="w-full border border-slate-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                    placeholder="Ej: Colegio San José, IES Begoña..." />
                </div>

                <label className={`block border-2 border-dashed rounded-xl p-8 text-center transition-colors ${importSourceName.trim() ? "border-indigo-300 hover:border-indigo-500 cursor-pointer" : "border-slate-200 opacity-50 cursor-not-allowed"}`}>
                  <div className="text-4xl mb-3">📊</div>
                  <span className="text-indigo-600 font-semibold text-sm">
                    {importSourceName.trim() ? "Seleccionar archivo" : "Introduce primero el nombre del colegio"}
                  </span>
                  <p className="text-slate-400 text-xs mt-1">.xlsx · .xls · .csv</p>
                  <input type="file" accept=".xlsx,.xls,.csv,.ods" onChange={handleFileSelected} disabled={!importSourceName.trim()} className="hidden" />
                </label>
              </div>
            </div>

            <div className="bg-blue-50 border border-blue-200 rounded-xl p-4 text-xs text-blue-700">
              <strong>🔒 Garantía de integridad:</strong> al reimportar un colegio, los trabajadores se identifican por DNI o email. Sus códigos se conservan aunque hayas añadido filas nuevas al Excel. Nunca se borran datos de otros colegios.
            </div>
          </div>
        )}

        {/* ── IMPORT: paso 2 preview ── */}
        {adminTab === "upload" && importStep === "preview" && previewData && (
          <div className="pb-8 space-y-4">
            {importProgress && (
              <div className="bg-indigo-50 border border-indigo-200 rounded-xl p-4 flex items-center gap-3">
                <div className="w-4 h-4 border-2 border-indigo-300 border-t-indigo-600 rounded-full animate-spin shrink-0" />
                <p className="text-indigo-700 text-sm">{importProgress}</p>
              </div>
            )}
            <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-6">
              <div className="flex items-start justify-between mb-4">
                <div>
                  <h3 className="font-bold text-slate-900">Revisar columnas</h3>
                  <p className="text-sm text-slate-500 mt-0.5">
                    <span className="font-medium text-indigo-600">🏫 {importSourceName}</span> · {previewData.fileName} · {previewData.rows.length} filas · {previewData.columns.length} columnas
                  </p>
                </div>
                <button onClick={() => setImportStep("name")} className="text-xs text-slate-400 hover:text-slate-600 border border-slate-200 px-3 py-1.5 rounded-lg">← Cambiar</button>
              </div>

              <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 mb-4 text-xs text-amber-700">
                ⚠️ Desmarca las columnas que no quieras incluir. Las columnas con huecos son las que se pedirá rellenar al trabajador.
              </div>

              <div className="space-y-2 max-h-80 overflow-y-auto pr-1">
                {previewData.columns.map(col => {
                  const type = guessColumnType(col);
                  const included = columnConfig[col] !== false;
                  const blankPct = Math.round(previewData.rows.filter(r => isBlank(r[col])).length / previewData.rows.length * 100);
                  const sample = previewData.rows.slice(0, 3).map(r => r[col]).filter(Boolean);
                  return (
                    <div key={col} onClick={() => setColumnConfig(p => ({ ...p, [col]: !included }))}
                      className={`flex items-center gap-3 p-3 rounded-xl border cursor-pointer transition-all ${included ? "bg-white border-slate-200 hover:border-indigo-300" : "bg-slate-50 border-slate-100 opacity-50"}`}>
                      <div className={`w-5 h-5 rounded-md border-2 flex items-center justify-center shrink-0 ${included ? "bg-indigo-600 border-indigo-600" : "border-slate-300 bg-white"}`}>
                        {included && <svg className="w-3 h-3 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={3} d="M5 13l4 4L19 7" /></svg>}
                      </div>
                      <span className="text-lg shrink-0">{type.icon}</span>
                      <div className="flex-1 min-w-0">
                        <div className="flex items-center gap-2">
                          <span className="text-sm font-semibold text-slate-800">{col}</span>
                          <span className="text-xs text-slate-400 bg-slate-100 px-1.5 py-0.5 rounded">{type.label}</span>
                        </div>
                        <div className="text-xs text-slate-400 mt-0.5 truncate">Ej: {sample.slice(0, 2).join(" · ") || "—sin datos—"}</div>
                      </div>
                      <div className="shrink-0 text-right">
                        {blankPct > 0
                          ? <span className="text-xs text-amber-600 font-medium">{blankPct}% vacíos</span>
                          : <span className="text-xs text-emerald-600 font-medium">100% relleno</span>}
                      </div>
                    </div>
                  );
                })}
              </div>

              {/* Análisis de duplicados */}
              {importAnalysis && (
                <div className="mt-4 rounded-xl overflow-hidden border border-slate-200">
                  <div className="bg-slate-50 px-4 py-2.5 border-b border-slate-200">
                    <p className="text-xs font-semibold text-slate-600">Análisis previo — {previewData.rows.length} filas en el archivo</p>
                  </div>
                  <div className="divide-y divide-slate-100">
                    <div className="flex items-center gap-3 px-4 py-2.5">
                      <div className="w-6 h-6 bg-emerald-100 rounded-lg flex items-center justify-center shrink-0">
                        <svg className="w-3.5 h-3.5 text-emerald-600" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M12 4v16m8-8H4"/></svg>
                      </div>
                      <div className="flex-1">
                        <p className="text-sm font-semibold text-slate-800">{importAnalysis.toAdd.length} trabajadores nuevos</p>
                        <p className="text-xs text-slate-400">Se añadirán al sistema con un código nuevo</p>
                      </div>
                    </div>
                    <div className="flex items-center gap-3 px-4 py-2.5">
                      <div className="w-6 h-6 bg-blue-100 rounded-lg flex items-center justify-center shrink-0">
                        <svg className="w-3.5 h-3.5 text-blue-600" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg>
                      </div>
                      <div className="flex-1">
                        <p className="text-sm font-semibold text-slate-800">{importAnalysis.toUpdate.length} ya existen y se actualizarán</p>
                        <p className="text-xs text-slate-400">
                          {importAnalysis.toUpdate.length > 0
                            ? `Identificados por: ${[...new Set(importAnalysis.toUpdate.map(t => t.method))].join(", ")}`
                            : "Conservan su código y sus respuestas enviadas"}
                        </p>
                      </div>
                    </div>
                    {importAnalysis.duplicateWarnings.length > 0 && (
                      <div className="flex items-start gap-3 px-4 py-2.5 bg-amber-50">
                        <div className="w-6 h-6 bg-amber-100 rounded-lg flex items-center justify-center shrink-0 mt-0.5">
                          <svg className="w-3.5 h-3.5 text-amber-600" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/></svg>
                        </div>
                        <div className="flex-1">
                          <p className="text-sm font-semibold text-amber-800">{importAnalysis.duplicateWarnings.length} posibles duplicados</p>
                          <p className="text-xs text-amber-600">Estas filas coinciden con un trabajador ya procesado en esta importación. Se añadirán como nuevos — revisa el Excel por si hay filas repetidas.</p>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              )}

              <div className="mt-4 pt-4 border-t border-slate-100 flex items-center justify-between">
                <p className="text-xs text-slate-400">{Object.values(columnConfig).filter(Boolean).length} de {previewData.columns.length} columnas seleccionadas</p>
                {importError && (
                  <div className="mb-3 bg-red-50 border border-red-200 text-red-700 text-sm px-4 py-3 rounded-xl">
                    ❌ {importError}
                  </div>
                )}
                <button onClick={handleConfirmImport} disabled={importStep === "saving" || Object.values(columnConfig).filter(Boolean).length === 0}
                  className="bg-indigo-600 text-white px-6 py-2.5 rounded-xl text-sm font-semibold hover:bg-indigo-700 disabled:opacity-40 transition-colors flex items-center gap-2">
                  {importStep === "saving"
                    ? <><div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" /> Guardando...</>
                    : <>Confirmar e importar →</>}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ── EXPORT ── */}
        {adminTab === "export" && (
          <div className="pb-8 space-y-4">
            <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-6">
              <h3 className="font-bold text-slate-900 mb-1">Datos actualizados</h3>
              <p className="text-sm text-slate-500 mb-4">Todos los trabajadores con sus datos originales y las respuestas recibidas. Incluye columna de colegio.</p>
              <div className="grid grid-cols-2 gap-3">
                <button onClick={handleExportDataXlsx} disabled={!workers.length} className="flex flex-col items-center gap-2 bg-emerald-50 border-2 border-emerald-200 hover:border-emerald-400 hover:bg-emerald-100 disabled:opacity-40 disabled:cursor-not-allowed rounded-xl p-4 transition-colors">
                  <span className="text-2xl">📗</span>
                  <span className="text-xs font-semibold text-emerald-700">Excel (.xlsx)</span>
                  <span className="text-xs text-emerald-500">Recomendado</span>
                </button>
                <button onClick={handleExportDataCsv} disabled={!workers.length} className="flex flex-col items-center gap-2 bg-slate-50 border-2 border-slate-200 hover:border-slate-400 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed rounded-xl p-4 transition-colors">
                  <span className="text-2xl">📄</span>
                  <span className="text-xs font-semibold text-slate-600">CSV (.csv)</span>
                  <span className="text-xs text-slate-400">Compatible</span>
                </button>
              </div>
            </div>
            <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-6">
              <h3 className="font-bold text-slate-900 mb-1">Lista de códigos</h3>
              <p className="text-sm text-slate-500 mb-3">Nombre, email, código y colegio de cada trabajador. Para mail merge.</p>
              <div className="bg-slate-50 border border-slate-100 rounded-xl p-3 mb-4 text-xs text-slate-500 font-mono">Nombre · Email · Código · Colegio · Estado · Campos vacíos</div>
              <div className="grid grid-cols-2 gap-3">
                <button onClick={handleExportCodesXlsx} disabled={!workers.length} className="flex flex-col items-center gap-2 bg-indigo-50 border-2 border-indigo-200 hover:border-indigo-400 hover:bg-indigo-100 disabled:opacity-40 disabled:cursor-not-allowed rounded-xl p-4 transition-colors">
                  <span className="text-2xl">📗</span>
                  <span className="text-xs font-semibold text-indigo-700">Excel (.xlsx)</span>
                  <span className="text-xs text-indigo-400">Recomendado</span>
                </button>
                <button onClick={handleExportCodesCsv} disabled={!workers.length} className="flex flex-col items-center gap-2 bg-slate-50 border-2 border-slate-200 hover:border-slate-400 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed rounded-xl p-4 transition-colors">
                  <span className="text-2xl">📄</span>
                  <span className="text-xs font-semibold text-slate-600">CSV (.csv)</span>
                  <span className="text-xs text-slate-400">Compatible</span>
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ── SETTINGS ── */}
        {adminTab === "settings" && (
          <div className="pb-8 space-y-4">
            {/* Mi cuenta */}
            <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-6">
              <div className="flex items-center gap-3 mb-4">
                <div className="w-10 h-10 bg-indigo-100 rounded-xl flex items-center justify-center text-lg font-bold text-indigo-600">
                  {currentAdmin?.name?.charAt(0).toUpperCase()}
                </div>
                <div>
                  <p className="font-bold text-slate-900">{currentAdmin?.name}</p>
                  <p className="text-xs text-slate-400">{currentAdmin?.isSuperAdmin ? "Superadministrador" : "Administrador"}</p>
                </div>
              </div>
              <h3 className="font-semibold text-slate-700 text-sm mb-3">Cambiar mi contraseña</h3>
              <div className="flex gap-2">
                <div className="relative flex-1">
                  <input type={showNewPass ? "text" : "password"} value={newPassInput}
                    onChange={e => setNewPassInput(e.target.value)}
                    onKeyDown={e => e.key === "Enter" && changeAdminPass()}
                    className="w-full border border-slate-200 rounded-xl px-4 py-2.5 pr-11 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                    placeholder="Nueva contraseña..." />
                  <button type="button" onClick={() => setShowNewPass(v => !v)} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600">
                    {showNewPass ? <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858.908a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242M9.88 9.88l-3.29-3.29m7.532 7.532l3.29 3.29M3 3l3.59 3.59m0 0A9.953 9.953 0 0112 5c4.478 0 8.268 2.943 9.543 7a10.025 10.025 0 01-4.132 5.411m0 0L21 21" /></svg> : <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" /><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" /></svg>}
                  </button>
                </div>
                <button onClick={changeAdminPass} disabled={!newPassInput.trim()} className="bg-indigo-600 text-white px-4 py-2 rounded-xl text-sm font-semibold hover:bg-indigo-700 disabled:opacity-40 transition-colors">Guardar</button>
              </div>
            </div>

            {/* Gestión de cuentas — solo superadmin */}
            {currentAdmin?.isSuperAdmin && (
              <div className="bg-white rounded-2xl border border-slate-100 shadow-sm p-6">
                <h3 className="font-bold text-slate-900 mb-1">Gestión de administradores</h3>
                <p className="text-sm text-slate-500 mb-4">Crea cuentas para que cada responsable entre con sus propias credenciales.</p>

                {/* Lista de cuentas */}
                <div className="space-y-2 mb-4">
                  {adminAccounts.map(acc => (
                    <div key={acc.id} className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-xl px-4 py-3">
                      <div className="w-8 h-8 bg-indigo-100 rounded-lg flex items-center justify-center text-sm font-bold text-indigo-600 shrink-0">
                        {acc.name.charAt(0).toUpperCase()}
                      </div>
                      <div className="flex-1 min-w-0">
                        <p className="text-sm font-semibold text-slate-800">{acc.name}</p>
                        <p className="text-xs text-slate-400">{acc.isSuperAdmin ? "⭐ Superadmin" : "Administrador"} · Contraseña: <span className="font-mono">{acc.password}</span></p>
                      </div>
                      {acc.id !== currentAdmin?.id && !acc.isSuperAdmin && (
                        <button onClick={() => deleteAdminAccount(acc.id)}
                          className="text-xs text-red-500 hover:text-red-700 border border-red-200 hover:border-red-400 px-2.5 py-1.5 rounded-lg transition-colors shrink-0">
                          Eliminar
                        </button>
                      )}
                      {acc.id === currentAdmin?.id && <span className="text-xs text-indigo-500 font-medium shrink-0">Tú</span>}
                    </div>
                  ))}
                </div>

                {/* Crear nueva cuenta */}
                <div className="border-t border-slate-100 pt-4">
                  <p className="text-sm font-semibold text-slate-700 mb-3">Crear nueva cuenta</p>
                  <div className="space-y-2">
                    <input type="text" value={newAccountName} onChange={e => setNewAccountName(e.target.value)}
                      className="w-full border border-slate-200 rounded-xl px-4 py-2.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                      placeholder="Nombre de usuario (ej: María García)" />
                    <div className="flex gap-2">
                      <div className="relative flex-1">
                        <input type={showNewAccountPass ? "text" : "password"} value={newAccountPass} onChange={e => setNewAccountPass(e.target.value)}
                          className="w-full border border-slate-200 rounded-xl px-4 py-2.5 pr-11 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
                          placeholder="Contraseña..." />
                        <button type="button" onClick={() => setShowNewAccountPass(v => !v)} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600">
                          {showNewAccountPass ? <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858.908a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242M9.88 9.88l-3.29-3.29m7.532 7.532l3.29 3.29M3 3l3.59 3.59m0 0A9.953 9.953 0 0112 5c4.478 0 8.268 2.943 9.543 7a10.025 10.025 0 01-4.132 5.411m0 0L21 21" /></svg> : <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" /><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" /></svg>}
                        </button>
                      </div>
                      <button onClick={createAdminAccount} disabled={!newAccountName.trim() || !newAccountPass.trim()}
                        className="bg-emerald-600 text-white px-4 py-2 rounded-xl text-sm font-semibold hover:bg-emerald-700 disabled:opacity-40 transition-colors shrink-0">
                        Crear
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </div>

      {/* Modal mensaje */}
      {msgModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4" onClick={() => setMsgModal(null)}>
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-sm p-6" onClick={e => e.stopPropagation()}>
            <div className="flex items-center justify-between mb-4">
              <h3 className="font-bold text-slate-900 text-sm">Mensaje para {msgModal.name}</h3>
              <button onClick={() => setMsgModal(null)} className="text-slate-400 hover:text-slate-600">
                <svg className="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" /></svg>
              </button>
            </div>
            <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 mb-3 select-all cursor-text">
              <pre className="text-sm text-slate-700 whitespace-pre-wrap font-sans">{msgModal.msg}</pre>
            </div>
            <p className="text-xs text-slate-400 text-center">Haz clic sobre el texto para seleccionarlo · luego Ctrl+C</p>
          </div>
        </div>
      )}
    </div>
  );

  return null;
}

import "./style.css";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { saveAs } from "file-saver";

/**
 * ✅ 기능 추가
 * 1) GLOBAL/VARIABLE 위치 이동(수정)
 * 2) 프로젝트별 저장/불러오기(프로필) + JSON 내보내기/가져오기
 * 3) UI 개선
 * 4) ✅ VARIABLE 값: 엑셀 복사/붙여넣기 일괄 추가 (탭/쉼표/줄바꿈 모두 지원)
 */

// ====== localStorage keys ======
const LS_KEY = "dsgen_state_v1";
const LS_PROFILES_KEY = "dsgen_profiles_v1";

function loadStateRaw() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}
function saveStateRaw(state) {
  localStorage.setItem(LS_KEY, JSON.stringify(state));
}

// ====== profiles storage ======
function loadProfiles() {
  try {
    const raw = localStorage.getItem(LS_PROFILES_KEY);
    if (!raw) return {};
    const obj = JSON.parse(raw);
    return obj && typeof obj === "object" ? obj : {};
  } catch {
    return {};
  }
}
function saveProfiles(profiles) {
  localStorage.setItem(LS_PROFILES_KEY, JSON.stringify(profiles || {}));
}

function nowIso() {
  return new Date().toISOString();
}

// ====== runtime state ======
let templateArrayBuffer = null;

let sheetNames = [];
let selectedSheetName = "";
let selectedCellAddr = "";

// preview
let previewGrid = [];
let previewMerges = [];
let previewColWidths = [];
let previewRowHeights = [];

// display range
let previewMaxR = 120;
let previewMaxC = 40;

// used range
let currentSheetRange = { rows: 0, cols: 0 };

// split
let splitLeftPx = 0;

// GLOBAL: { [sheetName]: { [addr]: value } }
let globalEdits = {};

// VARIABLE mapping: { [fieldKey]: [{sheetName, addr}] }
let variableMappings = {};

// VARIABLE values: { [fieldKey]: string[] }
let variableValues = {};

// filename rules
let fileNamePrefix = "PROJECT_";
let fileNameSuffix = "_Datasheet";
let fileNameField = "ItemNo";

// UI extra: profile + move mode
let profiles = loadProfiles();
let selectedProfileName = "";
let moveMode = null; // { kind:'global'|'variable', sheetName, addr, key? }

// ====== state helpers ======
function buildStateObject() {
  return {
    globalEdits,
    variableMappings,
    variableValues,
    fileNamePrefix,
    fileNameSuffix,
    fileNameField,
    previewMaxR,
    previewMaxC,
    splitLeftPx,
  };
}
function applyStateObject(s) {
  globalEdits = s.globalEdits || {};
  variableMappings = s.variableMappings || {};
  variableValues = s.variableValues || {};
  fileNamePrefix = s.fileNamePrefix ?? fileNamePrefix;
  fileNameSuffix = s.fileNameSuffix ?? fileNameSuffix;
  fileNameField = s.fileNameField ?? fileNameField;
  previewMaxR = s.previewMaxR ?? previewMaxR;
  previewMaxC = s.previewMaxC ?? previewMaxC;
  splitLeftPx = s.splitLeftPx ?? splitLeftPx;
}
function saveState() {
  saveStateRaw(buildStateObject());
}

// ====== restore persisted state ======
const persisted = loadStateRaw();
if (persisted) applyStateObject(persisted);

// ====== utils ======
function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}
function a1(r, c) {
  return XLSX.utils.encode_cell({ r, c });
}
function sanitizeFileName(name) {
  return String(name)
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim();
}
function ensureField(fieldKey) {
  if (!variableMappings[fieldKey]) variableMappings[fieldKey] = [];
  if (!variableValues[fieldKey]) variableValues[fieldKey] = [];
}
function uniqMappingsForCell(sheetName, addr) {
  for (const [k, arr] of Object.entries(variableMappings)) {
    if (arr?.some((m) => m.sheetName === sheetName && m.addr === addr))
      return { type: "variable", key: k };
  }
  if (globalEdits?.[sheetName]?.[addr] !== undefined) return { type: "global" };
  return null;
}
function flattenGlobals() {
  const out = [];
  for (const [sheet, cells] of Object.entries(globalEdits || {})) {
    for (const [addr, value] of Object.entries(cells || {})) out.push({ sheet, addr, value });
  }
  out.sort((a, b) => (a.sheet !== b.sheet ? a.sheet.localeCompare(b.sheet) : a.addr.localeCompare(b.addr)));
  return out;
}
function flattenVariableMaps() {
  const out = [];
  for (const [key, arr] of Object.entries(variableMappings || {})) {
    for (const m of arr || []) out.push({ key, sheetName: m.sheetName, addr: m.addr });
  }
  out.sort((a, b) => {
    if (a.key !== b.key) return a.key.localeCompare(b.key);
    if (a.sheetName !== b.sheetName) return a.sheetName.localeCompare(b.sheetName);
    return a.addr.localeCompare(b.addr);
  });
  return out;
}
function tryParseJson(text) {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

// ✅ 엑셀 복붙 텍스트를 값 리스트로 파싱
// - 기본: 탭/쉼표/세미콜론/줄바꿈을 모두 분리자로 처리
// - 여러 열을 복사한 경우: 모든 셀 값을 1차원으로 풀어서 추가
function parseBulkValues(text) {
  const raw = String(text ?? "");
  if (!raw.trim()) return [];

  // \r\n / \n 통일
  const normalized = raw.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

  // 엑셀 복사: 행은 \n, 열은 \t
  // 추가로 , ; 도 분리자로 지원
  const parts = normalized.split(/[\n\t,;]+/g);

  return parts
    .map((s) => String(s).trim())
    .filter((s) => s.length > 0);
}

// ✅ ws["!ref"] 기준으로 “빈칸 포함” 2D grid 생성
function buildGridFromRef(ws) {
  const ref = ws?.["!ref"];
  if (!ref) return { grid: [], range: { rows: 0, cols: 0 } };

  const rng = XLSX.utils.decode_range(ref);
  const rows = rng.e.r - rng.s.r + 1;
  const cols = rng.e.c - rng.s.c + 1;

  const grid = Array.from({ length: rows }, (_, rr) =>
    Array.from({ length: cols }, (_, cc) => {
      const r = rng.s.r + rr;
      const c = rng.s.c + cc;
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      return cell?.v ?? "";
    })
  );

  return { grid, range: { rows, cols } };
}

// ====== move logic ======
function startMoveGlobal(sheet, addr) {
  if (!globalEdits?.[sheet] || globalEdits[sheet][addr] === undefined) return;
  moveMode = { kind: "global", sheetName: sheet, addr };
  render();
}
function startMoveVariable(key, sheet, addr) {
  if (!variableMappings?.[key]) return;
  if (!variableMappings[key].some((m) => m.sheetName === sheet && m.addr === addr)) return;
  moveMode = { kind: "variable", key, sheetName: sheet, addr };
  render();
}
function cancelMove() {
  moveMode = null;
  render();
}
function applyMoveToTarget(targetSheet, targetAddr) {
  if (!moveMode) return;

  if (moveMode.sheetName === targetSheet && moveMode.addr === targetAddr) {
    moveMode = null;
    render();
    return;
  }

  if (moveMode.kind === "global") {
    const { sheetName: srcSheet, addr: srcAddr } = moveMode;
    const value = globalEdits?.[srcSheet]?.[srcAddr];
    if (value === undefined) {
      moveMode = null;
      return render();
    }

    delete globalEdits[srcSheet][srcAddr];
    if (Object.keys(globalEdits[srcSheet]).length === 0) delete globalEdits[srcSheet];

    if (!globalEdits[targetSheet]) globalEdits[targetSheet] = {};
    globalEdits[targetSheet][targetAddr] = value;

    saveState();
    moveMode = null;
    render();
    return;
  }

  if (moveMode.kind === "variable") {
    const { key, sheetName: srcSheet, addr: srcAddr } = moveMode;
    ensureField(key);

    variableMappings[key] = (variableMappings[key] || []).filter(
      (m) => !(m.sheetName === srcSheet && m.addr === srcAddr)
    );

    const exists = variableMappings[key].some((m) => m.sheetName === targetSheet && m.addr === targetAddr);
    if (!exists) variableMappings[key].push({ sheetName: targetSheet, addr: targetAddr });

    saveState();
    moveMode = null;
    render();
    return;
  }
}

// ====== UI render ======
function renderGrid(grid2d, merges = [], cols = [], rows = []) {
  if (!grid2d || grid2d.length === 0) {
    return `<div class="muted" style="padding:12px;">엑셀 업로드 후 시트를 선택하고 '시트 로드'를 누르세요.</div>`;
  }

  const maxR = Math.min(previewMaxR, grid2d.length);
  const maxC = Math.min(previewMaxC, Math.max(...grid2d.slice(0, maxR).map((r) => r.length || 0), 1));

  const mergeStart = new Map();
  const covered = new Set();

  for (const m of merges || []) {
    const r0 = m?.s?.r;
    const c0 = m?.s?.c;
    const r1 = m?.e?.r;
    const c1 = m?.e?.c;
    if ([r0, c0, r1, c1].some((x) => typeof x !== "number")) continue;

    mergeStart.set(`${r0},${c0}`, { rowspan: r1 - r0 + 1, colspan: c1 - c0 + 1 });

    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        if (r === r0 && c === c0) continue;
        covered.add(`${r},${c}`);
      }
    }
  }

  let colgroup = "";
  if (cols && cols.length) {
    colgroup = `<colgroup>${Array.from({ length: maxC })
      .map((_, c) => {
        const wpx = cols?.[c]?.wpx;
        const wch = cols?.[c]?.wch;
        const px = typeof wpx === "number" ? wpx : typeof wch === "number" ? Math.round(wch * 7) : null;
        return px ? `<col style="width:${px}px;" />` : `<col />`;
      })
      .join("")}</colgroup>`;
  }

  let html = `<table class="gridTable">${colgroup}`;

  for (let r = 0; r < maxR; r++) {
    const hpx = rows?.[r]?.hpx;
    const trStyle = typeof hpx === "number" ? ` style="height:${hpx}px;"` : "";
    html += `<tr${trStyle}>`;

    for (let c = 0; c < maxC; c++) {
      if (covered.has(`${r},${c}`)) continue;

      const addr = a1(r, c);
      const v = grid2d[r]?.[c] ?? "";
      const isSelected = addr === selectedCellAddr;

      const mark = uniqMappingsForCell(selectedSheetName, addr);
      let outline = "";
      if (isSelected) outline = "outline:2px solid #111; outline-offset:-2px;";
      else if (mark?.type === "global") outline = "outline:2px solid #16a34a; outline-offset:-2px;";
      else if (mark?.type === "variable") outline = "outline:2px solid #7c3aed; outline-offset:-2px;";

      const moveHint = moveMode ? "box-shadow: inset 0 0 0 1px rgba(0,0,0,0.06);" : "";

      const ms = mergeStart.get(`${r},${c}`);
      const rowspan = ms?.rowspan || 1;
      const colspan = ms?.colspan || 1;
      const cellExtraStyle = rowspan > 1 || colspan > 1 ? "vertical-align:middle;" : "";

      html += `
        <td data-cell="${addr}"
            title="${escapeHtml(selectedSheetName + "!" + addr)}"
            rowspan="${rowspan}"
            colspan="${colspan}"
            style="${cellExtraStyle} ${outline} ${moveHint}">
          ${escapeHtml(v)}
        </td>
      `;
    }

    html += `</tr>`;
  }

  html += `</table>`;
  return html;
}

function buildFileNamePreview() {
  const values = variableValues[fileNameField] || [];
  if (values.length === 0) return ["(파일명 필드 값이 아직 없어요)"];
  const preview = values.slice(0, 10).map((v) => {
    const core = v?.trim() ? v.trim() : "EMPTY";
    return sanitizeFileName(`${fileNamePrefix}${core}${fileNameSuffix}.xlsx`);
  });
  return values.length <= 10 ? preview : [...preview, `... (+${values.length - 10} more)`];
}

function renderVariableFields() {
  const keys = Object.keys(variableMappings);
  if (keys.length === 0) {
    return `<div class="emptyBox">VARIABLE로 등록된 필드가 아직 없습니다.</div>`;
  }

  return keys
    .map((key) => {
      const maps = variableMappings[key] || [];
      const vals = variableValues[key] || [];

      const mapBadges = maps.length
        ? maps
            .map(
              (m) => `
              <span class="pill pill-purple">
                ${escapeHtml(m.sheetName)}!${escapeHtml(m.addr)}
                <button class="pillBtn" data-movevar-key="${escapeHtml(key)}" data-movevar-sheet="${escapeHtml(
                m.sheetName
              )}" data-movevar-addr="${escapeHtml(m.addr)}">이동</button>
              </span>
            `
            )
            .join("")
        : `<span class="muted">(없음)</span>`;

      const listHtml = vals.length
        ? `<ol class="valueList">
            ${vals
              .map(
                (v, i) => `<li>
                  <span>${escapeHtml(v)}</span>
                  <button class="btn btn-ghost btn-sm" data-delval="${escapeHtml(key)}" data-idx="${i}">삭제</button>
                </li>`
              )
              .join("")}
          </ol>`
        : `<div class="muted" style="margin-top:8px;">아직 값이 없습니다.</div>`;

      return `
        <div class="card card-tight" style="margin-bottom:10px;">
          <div class="row row-between">
            <div>
              <div class="h3">${escapeHtml(key)}</div>
              <div class="muted" style="font-size:12px; margin-top:2px;">값 개수: <b>${vals.length}</b>개</div>
            </div>
            <div class="row" style="gap:6px; flex-wrap:wrap;">
              <button class="btn btn-danger btn-sm" data-clearfield="${escapeHtml(key)}">필드 삭제</button>
              <button class="btn btn-ghost btn-sm" data-clearvals="${escapeHtml(key)}">값 전부삭제</button>
            </div>
          </div>

          <div style="margin-top:8px;">
            <div class="label">매핑</div>
            <div class="wrap">${mapBadges}</div>
            <div class="muted" style="font-size:12px; margin-top:6px;">
              ※ [이동]을 누른 뒤, 프리뷰에서 새 셀을 클릭하면 매핑 위치가 바뀝니다.
            </div>
          </div>

          <div style="margin-top:10px;" class="row">
            <input class="input" data-addinput="${escapeHtml(key)}" placeholder="${escapeHtml(
        key
      )} 값 추가 (예: P-101)" style="flex:1; min-width:180px;" />
            <button class="btn btn-primary" data-addbtn="${escapeHtml(key)}">추가</button>
          </div>

          <!-- ✅ 일괄 붙여넣기 -->
          <div style="margin-top:10px;">
            <div class="label">일괄 붙여넣기</div>
            <textarea class="input" data-bulkarea="${escapeHtml(
              key
            )}" placeholder="엑셀에서 복사한 값을 여기에 붙여넣고 [일괄 추가]\n- 줄바꿈/탭/쉼표/세미콜론으로 자동 분리됩니다."
              style="width:100%; height:86px; resize:vertical; padding:10px; font-family:ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;"></textarea>
            <div class="row" style="margin-top:8px;">
              <button class="btn btn-primary btn-sm" data-bulkadd="${escapeHtml(key)}">일괄 추가</button>
              <button class="btn btn-ghost btn-sm" data-bulkclear="${escapeHtml(key)}">입력칸 비우기</button>
              <span class="muted" style="font-size:12px;">
                팁: 여러 열을 복사해도 모든 셀 값이 순서대로 추가돼요.
              </span>
            </div>
          </div>

          ${listHtml}
        </div>
      `;
    })
    .join("");
}

// ===== profiles UI helpers =====
function profileOptionsHtml() {
  const names = Object.keys(profiles || {}).sort((a, b) => a.localeCompare(b));
  if (names.length === 0) return `<option value="">(저장된 프로젝트 없음)</option>`;
  return [
    `<option value="">(선택)</option>`,
    ...names.map(
      (n) => `<option value="${escapeHtml(n)}" ${n === selectedProfileName ? "selected" : ""}>${escapeHtml(n)}</option>`
    ),
  ].join("");
}

function render() {
  const app = document.querySelector("#app");

  const fileCount = (variableValues[fileNameField] || []).length;
  const filePreview = buildFileNamePreview().map((x) => `- ${x}`).join("\n");

  const globalsFlat = flattenGlobals();
  const globalCount = globalsFlat.length;
  const varMapsFlat = flattenVariableMaps();

  const globalListHtml =
    globalsFlat.length === 0
      ? `<div class="muted" style="margin-top:8px;">등록된 GLOBAL이 없습니다.</div>`
      : `
        <div class="tableWrap" style="margin-top:8px;">
          <table class="table">
            <thead>
              <tr>
                <th style="width:44%;">위치</th>
                <th>값</th>
                <th style="width:160px;">동작</th>
              </tr>
            </thead>
            <tbody>
              ${globalsFlat
                .map(
                  (g) => `
                <tr>
                  <td><code>${escapeHtml(g.sheet)}!${escapeHtml(g.addr)}</code></td>
                  <td>${escapeHtml(g.value)}</td>
                  <td>
                    <div class="row" style="gap:6px; flex-wrap:wrap;">
                      <button class="btn btn-ghost btn-sm" data-moveglobal-sheet="${escapeHtml(
                        g.sheet
                      )}" data-moveglobal-addr="${escapeHtml(g.addr)}">이동</button>
                      <button class="btn btn-danger btn-sm" data-delglobal-sheet="${escapeHtml(
                        g.sheet
                      )}" data-delglobal-addr="${escapeHtml(g.addr)}">삭제</button>
                    </div>
                  </td>
                </tr>
              `
                )
                .join("")}
            </tbody>
          </table>
        </div>
      `;

  const varMapSummaryHtml =
    varMapsFlat.length === 0
      ? `<div class="muted" style="margin-top:6px;">VARIABLE 매핑이 없습니다.</div>`
      : `
        <div class="tableWrap" style="margin-top:8px;">
          <table class="table">
            <thead>
              <tr>
                <th style="width:26%;">필드</th>
                <th>위치</th>
                <th style="width:120px;">동작</th>
              </tr>
            </thead>
            <tbody>
              ${varMapsFlat
                .map(
                  (m) => `
                <tr>
                  <td><span class="pill pill-purple">${escapeHtml(m.key)}</span></td>
                  <td><code>${escapeHtml(m.sheetName)}!${escapeHtml(m.addr)}</code></td>
                  <td>
                    <button class="btn btn-ghost btn-sm"
                      data-movevar-key="${escapeHtml(m.key)}"
                      data-movevar-sheet="${escapeHtml(m.sheetName)}"
                      data-movevar-addr="${escapeHtml(m.addr)}"
                    >이동</button>
                  </td>
                </tr>
              `
                )
                .join("")}
            </tbody>
          </table>
        </div>
      `;

  const leftWidthStyle = splitLeftPx
    ? `grid-template-columns: ${splitLeftPx}px 10px 1fr;`
    : `grid-template-columns: 1.6fr 10px 1fr;`;

  const moveBannerHtml = moveMode
    ? `
      <div class="banner">
        <div>
          <b>이동 모드</b> :
          ${
            moveMode.kind === "global"
              ? `GLOBAL <code>${escapeHtml(moveMode.sheetName)}!${escapeHtml(moveMode.addr)}</code>`
              : `VARIABLE <span class="pill pill-purple">${escapeHtml(moveMode.key)}</span> <code>${escapeHtml(
                  moveMode.sheetName
                )}!${escapeHtml(moveMode.addr)}</code>`
          }
          을(를) 새 셀로 이동합니다. <span class="muted">프리뷰에서 원하는 셀을 클릭하세요.</span>
        </div>
        <div class="row" style="gap:8px;">
          <button class="btn btn-danger btn-sm" id="cancelMoveBtn">이동 취소</button>
        </div>
      </div>
    `
    : "";

  app.innerHTML = `
  <style>
    .wrapApp{ color:#111; font-family:system-ui, -apple-system, Segoe UI, Roboto, Arial; }
    .topbar{
      display:flex; align-items:flex-start; justify-content:space-between; gap:12px;
      padding:14px 16px; border:1px solid #e7e7e7; border-radius:16px;
      background:#ffffff; box-shadow:0 6px 18px rgba(0,0,0,.06);
      margin-bottom:12px;
    }
    .title{ margin:0; font-size:18px; font-weight:800; letter-spacing:-.2px; }
    .sub{ margin-top:4px; color:#666; font-size:12px; line-height:1.35; }
    .banner{
      display:flex; justify-content:space-between; gap:10px; align-items:center;
      padding:10px 12px; border-radius:14px; border:1px solid #fde68a;
      background:#fffbeb; color:#111; margin:10px 0 14px 0;
    }
    .card{
      border:1px solid #e7e7e7; border-radius:16px; padding:12px;
      background:#fff; box-shadow:0 6px 18px rgba(0,0,0,.06);
    }
    .card-tight{ padding:10px; box-shadow:none; }
    .h2{ margin:0 0 8px 0; font-size:14px; font-weight:800; }
    .h3{ font-size:13px; font-weight:800; }
    .label{ font-size:12px; color:#555; font-weight:700; margin-bottom:6px; }
    .muted{ color:#666; }
    .row{ display:flex; align-items:center; gap:8px; flex-wrap:wrap; }
    .row-between{ justify-content:space-between; }
    .input, select, textarea{
      border:1px solid #e5e7eb; border-radius:12px; padding:9px 10px; outline:none;
      font-size:13px; background:#fff;
    }
    .input:focus, select:focus, textarea:focus{ border-color:#93c5fd; box-shadow:0 0 0 3px rgba(59,130,246,.15); }
    .btn{
      border:1px solid transparent; border-radius:12px; padding:9px 12px;
      font-weight:800; font-size:13px; cursor:pointer;
      transition: transform .02s ease, box-shadow .15s ease, background .15s ease;
      user-select:none;
    }
    .btn:active{ transform: translateY(1px); }
    .btn-sm{ padding:7px 10px; font-size:12px; border-radius:11px; }
    .btn-primary{ background:#2563eb; color:white; box-shadow:0 6px 14px rgba(37,99,235,.20); }
    .btn-primary:hover{ background:#1d4ed8; }
    .btn-ghost{ background:#f3f4f6; color:#111; border-color:#e5e7eb; }
    .btn-ghost:hover{ background:#e5e7eb; }
    .btn-danger{ background:#ef4444; color:white; box-shadow:0 6px 14px rgba(239,68,68,.18); }
    .btn-danger:hover{ background:#dc2626; }
    .btn-ok{ background:#16a34a; color:white; box-shadow:0 6px 14px rgba(22,163,74,.18); }
    .btn-ok:hover{ background:#15803d; }
    .emptyBox{
      color:#666; padding:10px; border:1px dashed #ddd; border-radius:14px; background:#fafafa;
    }
    .pill{
      display:inline-flex; align-items:center; gap:6px;
      border-radius:999px; padding:6px 10px; font-size:12px; font-weight:800;
      border:1px solid #e5e7eb; background:#f9fafb; color:#111;
      margin-right:6px; margin-bottom:6px;
    }
    .pill-purple{ background:#f5f3ff; border-color:#ddd6fe; color:#4c1d95; }
    .pillBtn{
      border:0; background:transparent; cursor:pointer; font-weight:900; color:#111;
      padding:0 4px;
    }
    .pillBtn:hover{ text-decoration:underline; }
    .tableWrap{ border:1px solid #eee; border-radius:14px; overflow:auto; }
    .table{ width:100%; border-collapse:collapse; font-size:12px; }
    .table th{
      text-align:left; background:#fafafa; border-bottom:1px solid #eee; padding:10px;
      position: sticky; top: 0;
    }
    .table td{ padding:10px; border-bottom:1px solid #f2f2f2; vertical-align:top; }
    .valueList{ margin:10px 0 0 18px; padding:0; max-height:160px; overflow:auto; }
    .valueList li{ margin:4px 0; display:flex; justify-content:space-between; gap:8px; }
    .splitGrid{ display:grid; gap:16px; align-items:stretch; }
    .resizer{
      cursor:col-resize; border-radius:12px; background:#f3f4f6; border:1px solid #e5e7eb;
    }
    .gridTable{
      border-collapse:collapse; width:max-content; table-layout:fixed;
      font-size:12px;
    }
    .gridTable td{
      border:1px solid #eee; padding:4px; cursor:pointer;
      overflow:hidden; white-space:nowrap; text-overflow:ellipsis;
    }
    .sectionHead{
      display:flex; justify-content:space-between; align-items:flex-start; gap:10px;
    }
    .divider{ height:1px; background:#eee; margin:12px 0; }
  </style>

  <div class="wrapApp" style="padding:16px;">
    <div class="topbar">
      <div>
        <h1 class="title">Excel Datasheet Generator (ZIP)</h1>
        <div class="sub">프로젝트(프로필) 저장/불러오기 + 위치 이동 + 엑셀 붙여넣기 일괄 입력</div>
      </div>
      <div class="card" style="padding:10px; box-shadow:none; border-radius:14px; max-width:520px;">
        <div class="label">프로젝트 설정(프로필)</div>
        <div class="muted" style="font-size:12px;">현재 상태를 프로젝트별로 저장/불러오기 할 수 있어요.</div>

        <div class="row" style="margin-top:8px;">
          <select id="profileSelect" style="min-width:220px;">
            ${profileOptionsHtml()}
          </select>
          <button class="btn btn-ghost btn-sm" id="loadProfileBtn">불러오기</button>
          <button class="btn btn-danger btn-sm" id="deleteProfileBtn">삭제</button>
        </div>

        <div class="row" style="margin-top:8px;">
          <input class="input" id="profileNameInput" placeholder="새 프로젝트 이름 (예: 고객사A_2026Q1)" style="flex:1; min-width:220px;" />
          <button class="btn btn-primary btn-sm" id="saveNewProfileBtn">새로 저장</button>
          <button class="btn btn-ghost btn-sm" id="overwriteProfileBtn">덮어쓰기</button>
        </div>

        <div class="row" style="margin-top:8px;">
          <button class="btn btn-ghost btn-sm" id="exportProfileBtn">내보내기(JSON)</button>
          <label class="btn btn-ghost btn-sm" style="display:inline-flex; align-items:center; gap:8px;">
            가져오기(JSON)
            <input id="importProfileFile" type="file" accept="application/json,.json" style="display:none;" />
          </label>
        </div>
      </div>
    </div>

    ${moveBannerHtml}

    <div id="splitGrid" class="splitGrid" style="${leftWidthStyle}">
      <!-- LEFT -->
      <div id="leftPane" class="card" style="min-width:320px; overflow:hidden;">
        <div class="h2">1) 템플릿 업로드 & 시트 선택</div>
        <input id="fileInput" type="file" accept=".xlsx" />

        <div class="row" style="margin-top:10px;">
          <label class="label" style="margin:0;">Sheet</label>
          <select id="sheetSelect" ${sheetNames.length ? "" : "disabled"}>
            ${sheetNames
              .map(
                (n) =>
                  `<option value="${escapeHtml(n)}" ${n === selectedSheetName ? "selected" : ""}>${escapeHtml(n)}</option>`
              )
              .join("")}
          </select>
          <button class="btn btn-primary btn-sm" id="loadSheetBtn" ${selectedSheetName ? "" : "disabled"}>시트 로드</button>
          <div style="flex:1;"></div>
          <button class="btn btn-ghost btn-sm" id="resetSplitBtn" title="좌/우 폭 기본값으로">폭 초기화</button>
        </div>

        <div class="row" style="margin-top:10px;">
          <span class="muted" style="font-size:12px;">표시 범위</span>
          <label class="muted" style="font-size:12px;">행 <input id="maxRInput" type="number" min="10" max="2000" value="${previewMaxR}" style="width:90px;" /></label>
          <label class="muted" style="font-size:12px;">열 <input id="maxCInput" type="number" min="5" max="500" value="${previewMaxC}" style="width:90px;" /></label>
          <button class="btn btn-ghost btn-sm" id="applyPreviewRangeBtn">적용</button>
          <button class="btn btn-ghost btn-sm" id="fitUsedRangeBtn" ${currentSheetRange.rows ? "" : "disabled"} title="사용영역 전체로 맞춤">
            사용영역 전체 (${currentSheetRange.rows || 0}x${currentSheetRange.cols || 0})
          </button>
        </div>

        <div id="gridWrap" style="margin-top:10px; max-height:580px; overflow:auto; border:1px solid #eee; border-radius:14px;">
          ${renderGrid(previewGrid, previewMerges, previewColWidths, previewRowHeights)}
        </div>
      </div>

      <!-- RESIZER -->
      <div id="resizer" class="resizer" title="드래그해서 좌우 폭 조절"></div>

      <!-- RIGHT -->
      <div id="rightPane" class="card" style="min-width:320px;">
        <div class="h2">2) 셀 편집 등록</div>

        <div class="card-tight" style="border:1px solid #eee; border-radius:14px; background:#fafafa;">
          <div class="row">
            <span class="pill">${escapeHtml(selectedSheetName || "-")}</span>
            <span class="pill">${escapeHtml(selectedCellAddr || "-")}</span>
          </div>

          <div class="divider"></div>

          <div>
            <div class="label">GLOBAL</div>
            <div class="row">
              <input class="input" id="globalValueInput" placeholder="공통 문구 입력" style="flex:1; min-width:200px;" />
              <button class="btn btn-ok" id="applyGlobalBtn" ${selectedSheetName && selectedCellAddr ? "" : "disabled"}>GLOBAL 등록</button>
              <button class="btn btn-ghost" id="removeGlobalBtn" ${selectedSheetName && selectedCellAddr ? "" : "disabled"}>GLOBAL 해제</button>
            </div>
          </div>

          <div class="divider"></div>

          <div>
            <div class="label">VARIABLE</div>
            <div class="row">
              <input class="input" id="fieldKeyInput" placeholder="변수 필드명 (예: ItemNo)" style="flex:1; min-width:200px;" />
              <button class="btn btn-primary" id="applyVariableBtn" ${selectedSheetName && selectedCellAddr ? "" : "disabled"}>VARIABLE 등록</button>
              <button class="btn btn-ghost" id="removeVariableBtn" ${selectedSheetName && selectedCellAddr ? "" : "disabled"}>VARIABLE 해제</button>
            </div>
          </div>
        </div>

        <div style="margin-top:12px;">
          <div class="h2">GLOBAL 요약 <span class="muted" style="font-size:12px;">(총 ${globalCount}개)</span></div>
          ${globalListHtml}
        </div>

        <div style="margin-top:12px;">
          <div class="h2">VARIABLE 매핑 요약</div>
          ${varMapSummaryHtml}
        </div>

        <div class="divider"></div>

        <div>
          <div class="h2">3) VARIABLE 값 입력</div>
          ${renderVariableFields()}
        </div>

        <div class="divider"></div>

        <div>
          <div class="h2">4) 파일명 규칙 & Generate</div>

          <div class="card-tight" style="border:1px solid #eee; border-radius:14px;">
            <div class="row">
              <label class="label" style="margin:0;">Prefix</label>
              <input class="input" id="prefixInput" value="${escapeHtml(fileNamePrefix)}" style="flex:1; min-width:160px;" />
            </div>

            <div class="row" style="margin-top:8px;">
              <label class="label" style="margin:0;">파일명 필드</label>
              <select id="fileNameFieldSelect">
                ${
                  Object.keys(variableMappings).length
                    ? Object.keys(variableMappings)
                        .map(
                          (k) =>
                            `<option value="${escapeHtml(k)}" ${k === fileNameField ? "selected" : ""}>${escapeHtml(k)}</option>`
                        )
                        .join("")
                    : `<option value="${escapeHtml(fileNameField)}">${escapeHtml(fileNameField)}</option>`
                }
              </select>
              <span class="muted" style="font-size:12px;">(값 개수 = 생성 파일 개수)</span>
            </div>

            <div class="row" style="margin-top:8px;">
              <label class="label" style="margin:0;">Suffix</label>
              <input class="input" id="suffixInput" value="${escapeHtml(fileNameSuffix)}" style="flex:1; min-width:160px;" />
            </div>

            <div style="margin-top:8px;">생성 파일 개수: <b>${fileCount}</b></div>

            <details style="margin-top:8px;">
              <summary style="cursor:pointer; font-weight:800;">파일명 미리보기(최대 10개)</summary>
              <pre style="background:#fff; border:1px solid #eee; padding:10px; border-radius:14px; max-height:160px; overflow:auto;">${escapeHtml(
                filePreview
              )}</pre>
            </details>

            <div class="row" style="margin-top:10px;">
              <button class="btn btn-primary" id="generateZipBtn">Generate ZIP</button>
              <button class="btn btn-danger" id="resetAllBtn">전체 초기화(저장 포함)</button>
            </div>
          </div>
        </div>

      </div>
    </div>
  </div>
  `;

  // ===== events =====

  // move banner cancel
  document.querySelector("#cancelMoveBtn")?.addEventListener("click", cancelMove);

  // profile events
  document.querySelector("#profileSelect")?.addEventListener("change", (e) => {
    selectedProfileName = e.target.value || "";
  });

  document.querySelector("#loadProfileBtn")?.addEventListener("click", () => {
    if (!selectedProfileName) return alert("불러올 프로젝트를 선택하세요.");
    const p = profiles[selectedProfileName];
    if (!p) return alert("프로젝트를 찾지 못했어요.");
    if (!confirm(`'${selectedProfileName}' 프로젝트 설정을 불러올까요? (현재 설정은 덮어써짐)`)) return;

    applyStateObject(p.state || {});
    saveState();
    render();
  });

  document.querySelector("#deleteProfileBtn")?.addEventListener("click", () => {
    if (!selectedProfileName) return alert("삭제할 프로젝트를 선택하세요.");
    if (!confirm(`'${selectedProfileName}' 프로젝트를 삭제할까요?`)) return;
    delete profiles[selectedProfileName];
    saveProfiles(profiles);
    selectedProfileName = "";
    render();
  });

  document.querySelector("#saveNewProfileBtn")?.addEventListener("click", () => {
    const name = (document.querySelector("#profileNameInput")?.value ?? "").trim();
    if (!name) return alert("프로젝트 이름을 입력하세요.");
    if (profiles[name]) return alert("같은 이름의 프로젝트가 이미 있어요. (덮어쓰기 사용)");
    profiles[name] = { name, updatedAt: nowIso(), state: buildStateObject() };
    saveProfiles(profiles);
    selectedProfileName = name;
    alert(`저장 완료: ${name}`);
    render();
  });

  document.querySelector("#overwriteProfileBtn")?.addEventListener("click", () => {
    const name = selectedProfileName || (document.querySelector("#profileNameInput")?.value ?? "").trim();
    if (!name) return alert("덮어쓸 프로젝트를 선택하거나 이름을 입력하세요.");
    if (!confirm(`'${name}' 프로젝트에 현재 설정을 덮어쓸까요?`)) return;
    profiles[name] = { name, updatedAt: nowIso(), state: buildStateObject() };
    saveProfiles(profiles);
    selectedProfileName = name;
    alert(`덮어쓰기 완료: ${name}`);
    render();
  });

  document.querySelector("#exportProfileBtn")?.addEventListener("click", () => {
    const name = selectedProfileName;
    if (!name) return alert("내보낼 프로젝트를 선택하세요.");
    const p = profiles[name];
    if (!p) return alert("프로젝트를 찾지 못했어요.");
    const blob = new Blob([JSON.stringify(p, null, 2)], { type: "application/json" });
    saveAs(blob, sanitizeFileName(`${name}_profile.json`));
  });

  document.querySelector("#importProfileFile")?.addEventListener("change", async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const obj = tryParseJson(text);
      if (!obj || !obj.name || !obj.state) throw new Error("올바른 프로필 JSON 형식이 아닙니다.");
      const name = String(obj.name).trim();
      if (!name) throw new Error("프로필 name이 비어있습니다.");

      if (profiles[name]) {
        const ok = confirm(`'${name}' 프로필이 이미 있어요. 덮어쓸까요?`);
        if (!ok) return;
      }
      profiles[name] = { name, updatedAt: nowIso(), state: obj.state };
      saveProfiles(profiles);
      selectedProfileName = name;
      alert(`가져오기 완료: ${name}`);
      render();
    } catch (err) {
      alert("가져오기 실패: " + (err?.message || String(err)));
    } finally {
      e.target.value = "";
    }
  });

  // split reset
  document.querySelector("#resetSplitBtn")?.addEventListener("click", () => {
    splitLeftPx = 0;
    saveState();
    render();
  });

  // file / sheet
  document.querySelector("#fileInput").addEventListener("change", onFile);

  document.querySelector("#sheetSelect")?.addEventListener("change", (e) => {
    selectedSheetName = e.target.value;
    selectedCellAddr = "";
    previewGrid = [];
    previewMerges = [];
    previewColWidths = [];
    previewRowHeights = [];
    currentSheetRange = { rows: 0, cols: 0 };
    render();
  });

  document.querySelector("#loadSheetBtn")?.addEventListener("click", onLoadSelectedSheet);

  document.querySelector("#applyPreviewRangeBtn")?.addEventListener("click", () => {
    const r = Number(document.querySelector("#maxRInput")?.value || previewMaxR);
    const c = Number(document.querySelector("#maxCInput")?.value || previewMaxC);
    previewMaxR = Math.max(10, Math.min(2000, r));
    previewMaxC = Math.max(5, Math.min(500, c));
    saveState();
    render();
  });

  document.querySelector("#fitUsedRangeBtn")?.addEventListener("click", () => {
    if (!currentSheetRange.rows || !currentSheetRange.cols) return;
    previewMaxR = Math.min(2000, currentSheetRange.rows);
    previewMaxC = Math.min(500, currentSheetRange.cols);
    saveState();
    render();
  });

  // grid click: select cell OR apply move
  document.querySelector("#gridWrap")?.addEventListener("click", (ev) => {
    const td = ev.target?.closest?.("td[data-cell]");
    if (!td) return;
    const addr = td.getAttribute("data-cell") || "";
    selectedCellAddr = addr;

    if (moveMode) {
      if (!selectedSheetName) return render();
      applyMoveToTarget(selectedSheetName, addr);
      return;
    }

    render();
  });

  // global apply/remove
  document.querySelector("#applyGlobalBtn")?.addEventListener("click", onApplyGlobal);
  document.querySelector("#removeGlobalBtn")?.addEventListener("click", onRemoveGlobal);

  document.querySelectorAll("[data-delglobal-sheet]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const sheet = btn.getAttribute("data-delglobal-sheet");
      const addr = btn.getAttribute("data-delglobal-addr");
      if (!sheet || !addr) return;
      if (globalEdits?.[sheet]) {
        delete globalEdits[sheet][addr];
        if (Object.keys(globalEdits[sheet]).length === 0) delete globalEdits[sheet];
      }
      saveState();
      render();
    });
  });

  document.querySelectorAll("[data-moveglobal-sheet]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const sheet = btn.getAttribute("data-moveglobal-sheet");
      const addr = btn.getAttribute("data-moveglobal-addr");
      if (!sheet || !addr) return;
      startMoveGlobal(sheet, addr);
    });
  });

  // variable apply/remove
  document.querySelector("#applyVariableBtn")?.addEventListener("click", onApplyVariable);
  document.querySelector("#removeVariableBtn")?.addEventListener("click", onRemoveVariable);

  // move variable mapping buttons
  document.querySelectorAll("[data-movevar-key]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.getAttribute("data-movevar-key");
      const sheet = btn.getAttribute("data-movevar-sheet");
      const addr = btn.getAttribute("data-movevar-addr");
      if (!key || !sheet || !addr) return;
      startMoveVariable(key, sheet, addr);
    });
  });

  // variable field events (single add)
  document.querySelectorAll("[data-addbtn]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.getAttribute("data-addbtn");
      const input = document.querySelector(`[data-addinput="${CSS.escape(key)}"]`);
      const v = (input?.value ?? "").trim();
      if (!v) return;
      ensureField(key);
      variableValues[key].push(v);
      input.value = "";
      saveState();
      render();
    });
  });

  // ✅ bulk add
  document.querySelectorAll("[data-bulkadd]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.getAttribute("data-bulkadd");
      if (!key) return;
      const area = document.querySelector(`[data-bulkarea="${CSS.escape(key)}"]`);
      const text = area?.value ?? "";
      const values = parseBulkValues(text);

      if (values.length === 0) return alert("붙여넣을 값이 없습니다.");

      ensureField(key);
      // 그대로 추가 (중복 허용). 중복 제거 원하면 여기서 Set 처리 가능.
      variableValues[key].push(...values);

      area.value = "";
      saveState();
      render();
    });
  });

  document.querySelectorAll("[data-bulkclear]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.getAttribute("data-bulkclear");
      if (!key) return;
      const area = document.querySelector(`[data-bulkarea="${CSS.escape(key)}"]`);
      if (area) area.value = "";
    });
  });

  // delete one value
  document.querySelectorAll("[data-delval]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.getAttribute("data-delval");
      const idx = Number(btn.getAttribute("data-idx"));
      ensureField(key);
      variableValues[key].splice(idx, 1);
      saveState();
      render();
    });
  });

  // clear values
  document.querySelectorAll("[data-clearvals]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.getAttribute("data-clearvals");
      if (!confirm(`[${key}] 값 목록을 전부 삭제할까요?`)) return;
      ensureField(key);
      variableValues[key] = [];
      saveState();
      render();
    });
  });

  // clear field
  document.querySelectorAll("[data-clearfield]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.getAttribute("data-clearfield");
      if (!confirm(`[${key}] 필드(매핑+값)를 삭제할까요?`)) return;
      delete variableMappings[key];
      delete variableValues[key];
      if (fileNameField === key) fileNameField = "ItemNo";
      saveState();
      render();
    });
  });

  // filename rules
  document.querySelector("#prefixInput")?.addEventListener("input", (e) => {
    fileNamePrefix = e.target.value ?? "";
    saveState();
  });
  document.querySelector("#suffixInput")?.addEventListener("input", (e) => {
    fileNameSuffix = e.target.value ?? "";
    saveState();
  });
  document.querySelector("#fileNameFieldSelect")?.addEventListener("change", (e) => {
    fileNameField = e.target.value || "ItemNo";
    saveState();
    render();
  });

  // generate
  document.querySelector("#generateZipBtn")?.addEventListener("click", onGenerateZip);

  // resizer drag
  (function bindResizer() {
    const resizer = document.querySelector("#resizer");
    const splitGrid = document.querySelector("#splitGrid");
    if (!resizer || !splitGrid) return;

    let dragging = false;

    const onMove = (e) => {
      if (!dragging) return;

      const rect = splitGrid.getBoundingClientRect();
      let newLeft = e.clientX - rect.left;

      const MIN_LEFT = 320;
      const MIN_RIGHT = 320;
      const maxLeft = rect.width - MIN_RIGHT - 10;

      newLeft = Math.max(MIN_LEFT, Math.min(maxLeft, newLeft));

      splitLeftPx = Math.round(newLeft);
      saveState();

      splitGrid.style.gridTemplateColumns = `${splitLeftPx}px 10px 1fr`;
    };

    const stop = () => {
      dragging = false;
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
      window.removeEventListener("mousemove", onMove);
      window.removeEventListener("mouseup", stop);
    };

    resizer.addEventListener("mousedown", () => {
      dragging = true;
      document.body.style.cursor = "col-resize";
      document.body.style.userSelect = "none";
      window.addEventListener("mousemove", onMove);
      window.addEventListener("mouseup", stop);
    });
  })();

  // reset all
  document.querySelector("#resetAllBtn")?.addEventListener("click", () => {
    if (!confirm("정말 전체 초기화할까요? (저장된 값도 삭제됨)")) return;
    globalEdits = {};
    variableMappings = {};
    variableValues = {};
    fileNamePrefix = "PROJECT_";
    fileNameSuffix = "_Datasheet";
    fileNameField = "ItemNo";
    templateArrayBuffer = null;
    sheetNames = [];
    selectedSheetName = "";
    selectedCellAddr = "";
    previewGrid = [];
    previewMerges = [];
    previewColWidths = [];
    previewRowHeights = [];
    currentSheetRange = { rows: 0, cols: 0 };
    splitLeftPx = 0;
    moveMode = null;
    localStorage.removeItem(LS_KEY);
    render();
  });
}

// ===== handlers =====
async function onFile(e) {
  const file = e.target.files?.[0];
  if (!file) return;

  templateArrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(templateArrayBuffer, { type: "array" });

  sheetNames = wb.SheetNames || [];
  selectedSheetName = sheetNames[0] || "";
  selectedCellAddr = "";

  previewGrid = [];
  previewMerges = [];
  previewColWidths = [];
  previewRowHeights = [];
  currentSheetRange = { rows: 0, cols: 0 };

  alert(`업로드 완료! 시트 ${sheetNames.length}개를 찾았어요.`);
  render();
}

function onLoadSelectedSheet() {
  if (!templateArrayBuffer) return alert("먼저 xlsx 파일을 업로드하세요.");
  if (!selectedSheetName) return alert("시트를 선택하세요.");

  const wb = XLSX.read(templateArrayBuffer, { type: "array" });
  const ws = wb.Sheets[selectedSheetName];
  if (!ws) return alert(`시트 '${selectedSheetName}'를 찾지 못했어요.`);

  const { grid, range } = buildGridFromRef(ws);
  previewGrid = grid;
  currentSheetRange = range;

  previewMerges = ws["!merges"] || [];
  previewColWidths = ws["!cols"] || [];
  previewRowHeights = ws["!rows"] || [];

  selectedCellAddr = "";
  render();
}

function onApplyGlobal() {
  if (!selectedSheetName || !selectedCellAddr) return;
  const value = document.querySelector("#globalValueInput")?.value ?? "";
  if (!globalEdits[selectedSheetName]) globalEdits[selectedSheetName] = {};
  globalEdits[selectedSheetName][selectedCellAddr] = value;
  saveState();
  render();
}

function onRemoveGlobal() {
  if (!selectedSheetName || !selectedCellAddr) return;
  if (globalEdits[selectedSheetName]) {
    delete globalEdits[selectedSheetName][selectedCellAddr];
    if (Object.keys(globalEdits[selectedSheetName]).length === 0) delete globalEdits[selectedSheetName];
  }
  saveState();
  render();
}

function onApplyVariable() {
  if (!selectedSheetName || !selectedCellAddr) return;
  const key = (document.querySelector("#fieldKeyInput")?.value ?? "").trim();
  if (!key) return alert("변수 필드명을 입력하세요. 예: ItemNo, TagNo");

  ensureField(key);

  const exists = variableMappings[key].some((m) => m.sheetName === selectedSheetName && m.addr === selectedCellAddr);
  if (!exists) variableMappings[key].push({ sheetName: selectedSheetName, addr: selectedCellAddr });

  if (!variableMappings[fileNameField]) fileNameField = key;

  saveState();
  render();
}

function onRemoveVariable() {
  if (!selectedSheetName || !selectedCellAddr) return;

  let removedAny = false;
  for (const [key, arr] of Object.entries(variableMappings)) {
    const before = arr.length;
    variableMappings[key] = arr.filter((m) => !(m.sheetName === selectedSheetName && m.addr === selectedCellAddr));
    if (variableMappings[key].length !== before) removedAny = true;
    if (variableMappings[key].length === 0) {
      delete variableMappings[key];
      delete variableValues[key];
    }
  }

  if (!removedAny) alert("이 셀에 등록된 VARIABLE 매핑이 없어요.");
  if (!variableMappings[fileNameField]) fileNameField = "ItemNo";

  saveState();
  render();
}

async function buildOneFileXlsx(itemIndex) {
  const outWb = new ExcelJS.Workbook();
  await outWb.xlsx.load(templateArrayBuffer);

  for (const [sheetName, cells] of Object.entries(globalEdits)) {
    const s = outWb.getWorksheet(sheetName);
    if (!s) continue;
    for (const [addr, value] of Object.entries(cells)) s.getCell(addr).value = value;
  }

  for (const [fieldKey, maps] of Object.entries(variableMappings)) {
    const values = variableValues[fieldKey] || [];
    const v = values[itemIndex] ?? "";
    for (const m of maps) {
      const s = outWb.getWorksheet(m.sheetName);
      if (!s) continue;
      s.getCell(m.addr).value = v;
    }
  }

  return await outWb.xlsx.writeBuffer();
}

async function onGenerateZip() {
  try {
    if (!templateArrayBuffer) return alert("템플릿 xlsx를 업로드하세요.");

    if (!variableMappings[fileNameField] || (variableValues[fileNameField] || []).length === 0) {
      return alert(
        `파일명 필드(${fileNameField})에 값이 1개 이상 필요해요.\n현재 ${fileNameField} 값 개수: ${(variableValues[fileNameField] || []).length}`
      );
    }

    const values = variableValues[fileNameField] || [];
    alert(`파일 생성 개수: ${values.length}개 (ZIP 생성 중...)`);

    const zip = new JSZip();
    const zipName = sanitizeFileName(`${fileNamePrefix}${fileNameSuffix || ""}_OUTPUT.zip`) || "output.zip";

    for (let i = 0; i < values.length; i++) {
      const core = (values[i] ?? "").trim();
      const baseName = `${fileNamePrefix}${core || `DS_${String(i + 1).padStart(3, "0")}`}${fileNameSuffix}.xlsx`;
      const fileName = sanitizeFileName(baseName);

      const buf = await buildOneFileXlsx(i);
      zip.file(fileName, buf);
    }

    const zipBlob = await zip.generateAsync({ type: "blob" });
    saveAs(zipBlob, zipName);
  } catch (err) {
    console.error(err);
    alert("에러 발생: " + (err?.message || String(err)));
  }
}

// initial render
render();

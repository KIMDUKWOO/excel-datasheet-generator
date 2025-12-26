import "./style.css";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import JSZip from "jszip";
import { saveAs } from "file-saver";

/**
 * ✅ 개선 사항
 * - 프리뷰가 35행/20열로 잘려서 오른쪽/아래가 안 보이던 문제 해결
 *   -> previewMaxR/previewMaxC로 표시범위 조절 + "사용영역 전체" 버튼 제공
 * - sheet_to_json이 trailing blank를 잘라먹는 문제 해결
 *   -> ws["!ref"] 범위 기준으로 셀을 직접 읽어 2D grid 생성
 */

// ====== localStorage ======
const LS_KEY = "dsgen_state_v1";
function loadState() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch {
    return null;
  }
}
function saveState() {
  const state = {
    globalEdits,
    variableMappings,
    variableValues,
    fileNamePrefix,
    fileNameSuffix,
    fileNameField,
    previewMaxR,
    previewMaxC,
  };
  localStorage.setItem(LS_KEY, JSON.stringify(state));
}

// ====== runtime state ======
let templateArrayBuffer = null;

let sheetNames = [];
let selectedSheetName = "";
let selectedCellAddr = "";

// 프리뷰
let previewGrid = [];
let previewMerges = [];      // ws["!merges"]
let previewColWidths = [];   // ws["!cols"]
let previewRowHeights = [];  // ws["!rows"]

// ✅ 표시 범위(기본값 넉넉하게)
let previewMaxR = 120;
let previewMaxC = 40;

// ✅ 현재 시트의 실제 사용 영역(!ref) 정보
let currentSheetRange = { rows: 0, cols: 0 };

// GLOBAL: { [sheetName]: { [addr]: value } }
let globalEdits = {};

// VARIABLE mapping: { [fieldKey]: [{sheetName, addr}] }
let variableMappings = {};

// VARIABLE values: { [fieldKey]: string[] }
let variableValues = {};

// 파일명 규칙
let fileNamePrefix = "PROJECT_";
let fileNameSuffix = "_Datasheet";
let fileNameField = "ItemNo";

// ====== restore persisted state ======
const persisted = loadState();
if (persisted) {
  globalEdits = persisted.globalEdits || {};
  variableMappings = persisted.variableMappings || {};
  variableValues = persisted.variableValues || {};
  fileNamePrefix = persisted.fileNamePrefix ?? fileNamePrefix;
  fileNameSuffix = persisted.fileNameSuffix ?? fileNameSuffix;
  fileNameField = persisted.fileNameField ?? fileNameField;
  previewMaxR = persisted.previewMaxR ?? previewMaxR;
  previewMaxC = persisted.previewMaxC ?? previewMaxC;
}

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
function uniqMappingsForCell(sheetName, addr) {
  for (const [k, arr] of Object.entries(variableMappings)) {
    if (arr?.some((m) => m.sheetName === sheetName && m.addr === addr))
      return { type: "variable", key: k };
  }
  if (globalEdits?.[sheetName]?.[addr] !== undefined) return { type: "global" };
  return null;
}
function ensureField(fieldKey) {
  if (!variableMappings[fieldKey]) variableMappings[fieldKey] = [];
  if (!variableValues[fieldKey]) variableValues[fieldKey] = [];
}

function flattenGlobals() {
  const out = [];
  for (const [sheet, cells] of Object.entries(globalEdits || {})) {
    for (const [addr, value] of Object.entries(cells || {})) {
      out.push({ sheet, addr, value });
    }
  }
  out.sort((a, b) => (a.sheet !== b.sheet ? a.sheet.localeCompare(b.sheet) : a.addr.localeCompare(b.addr)));
  return out;
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
      return cell?.v ?? ""; // 빈칸 포함
    })
  );

  return { grid, range: { rows, cols } };
}

// ====== UI render ======
function renderGrid(grid2d, merges = [], cols = [], rows = []) {
  if (!grid2d || grid2d.length === 0) {
    return `<div style="padding:12px; color:#777;">엑셀 업로드 후 시트를 선택하고 '시트 로드'를 누르세요.</div>`;
  }

  const maxR = Math.min(previewMaxR, grid2d.length);
  const maxC = Math.min(previewMaxC, Math.max(...grid2d.slice(0, maxR).map((r) => r.length || 0), 1));

  // ---- merge maps ----
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

  // ---- column widths via <colgroup> (best-effort) ----
  let colgroup = "";
  if (cols && cols.length) {
    colgroup = `<colgroup>${Array.from({ length: maxC })
      .map((_, c) => {
        const wpx = cols?.[c]?.wpx;
        const wch = cols?.[c]?.wch;
        const px =
          typeof wpx === "number" ? wpx : typeof wch === "number" ? Math.round(wch * 7) : null;
        return px ? `<col style="width:${px}px;" />` : `<col />`;
      })
      .join("")}</colgroup>`;
  }

  let html = `<table style="border-collapse:collapse; width:max-content; table-layout:fixed;">${colgroup}`;

  for (let r = 0; r < maxR; r++) {
    const hpx = rows?.[r]?.hpx;
    const trStyle = typeof hpx === "number" ? ` style="height:${hpx}px;"` : "";
    html += `<tr${trStyle}>`;

    for (let c = 0; c < maxC; c++) {
      // merge covered는 skip
      if (covered.has(`${r},${c}`)) continue;

      const addr = a1(r, c);
      const v = grid2d[r]?.[c] ?? "";
      const isSelected = addr === selectedCellAddr;

      const mark = uniqMappingsForCell(selectedSheetName, addr);
      let outline = "";
      if (isSelected) outline = "outline:2px solid #000;";
      else if (mark?.type === "global") outline = "outline:2px solid #00a86b;";
      else if (mark?.type === "variable") outline = "outline:2px solid #7a5cff;";

      const ms = mergeStart.get(`${r},${c}`);
      const rowspan = ms?.rowspan || 1;
      const colspan = ms?.colspan || 1;
      const cellExtraStyle = rowspan > 1 || colspan > 1 ? "vertical-align:middle;" : "";

      html += `
        <td data-cell="${addr}"
            title="${escapeHtml(selectedSheetName + "!" + addr)}"
            rowspan="${rowspan}"
            colspan="${colspan}"
            style="border:1px solid #eee; padding:4px; font-size:12px; cursor:pointer;
                   overflow:hidden; white-space:nowrap; text-overflow:ellipsis; ${cellExtraStyle} ${outline}">
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
    return `<div style="color:#777; padding:8px; border:1px dashed #ddd; border-radius:8px;">VARIABLE로 등록된 필드가 아직 없습니다.</div>`;
  }

  return keys
    .map((key) => {
      const maps = variableMappings[key] || [];
      const vals = variableValues[key] || [];
      const mapText = maps.map((m) => `${m.sheetName}!${m.addr}`).join(", ");

      const listHtml = vals.length
        ? `<ol style="margin:8px 0 0 18px; padding:0; max-height:120px; overflow:auto;">
            ${vals
              .map(
                (v, i) => `<li style="margin:2px 0;">
                  ${escapeHtml(v)}
                  <button data-delval="${escapeHtml(key)}" data-idx="${i}" style="margin-left:8px;">삭제</button>
                </li>`
              )
              .join("")}
          </ol>`
        : `<div style="color:#777; margin-top:8px;">아직 값이 없습니다.</div>`;

      return `
        <div style="border:1px solid #eee; border-radius:10px; padding:10px; margin-bottom:10px;">
          <div style="display:flex; justify-content:space-between; gap:8px; align-items:flex-start;">
            <div>
              <div style="font-weight:700;">${escapeHtml(key)}</div>
              <div style="color:#666; font-size:12px; margin-top:2px;">매핑: ${escapeHtml(mapText || "(없음)")}</div>
              <div style="color:#666; font-size:12px; margin-top:2px;">값 개수: ${vals.length}개</div>
            </div>
            <div style="display:flex; gap:6px; flex-wrap:wrap;">
              <button data-clearfield="${escapeHtml(key)}">필드 삭제</button>
              <button data-clearvals="${escapeHtml(key)}">값 전부삭제</button>
            </div>
          </div>

          <div style="margin-top:8px; display:flex; gap:6px; flex-wrap:wrap; align-items:center;">
            <input data-addinput="${escapeHtml(key)}" placeholder="${escapeHtml(key)} 값 추가 (예: P-101)" style="flex:1; min-width:180px;" />
            <button data-addbtn="${escapeHtml(key)}">추가</button>
          </div>

          ${listHtml}
        </div>
      `;
    })
    .join("");
}

function render() {
  const app = document.querySelector("#app");

  const fileCount = (variableValues[fileNameField] || []).length;
  const filePreview = buildFileNamePreview().map((x) => `- ${x}`).join("\n");

  const globalsFlat = flattenGlobals();
  const globalCount = globalsFlat.length;

  const globalListHtml =
    globalsFlat.length === 0
      ? `<div style="color:#777; margin-top:8px;">등록된 GLOBAL이 없습니다.</div>`
      : `
        <div style="margin-top:8px; max-height:200px; overflow:auto; border:1px solid #eee; border-radius:10px;">
          <table style="width:100%; border-collapse:collapse; font-size:12px;">
            <thead>
              <tr style="background:#fafafa;">
                <th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:40%;">위치</th>
                <th style="text-align:left; padding:8px; border-bottom:1px solid #eee;">값</th>
                <th style="text-align:left; padding:8px; border-bottom:1px solid #eee; width:80px;">동작</th>
              </tr>
            </thead>
            <tbody>
              ${globalsFlat
                .map(
                  (g) => `
                <tr>
                  <td style="padding:8px; border-bottom:1px solid #f2f2f2;">
                    <code>${escapeHtml(g.sheet)}!${escapeHtml(g.addr)}</code>
                  </td>
                  <td style="padding:8px; border-bottom:1px solid #f2f2f2;">
                    ${escapeHtml(g.value)}
                  </td>
                  <td style="padding:8px; border-bottom:1px solid #f2f2f2;">
                    <button data-delglobal-sheet="${escapeHtml(g.sheet)}" data-delglobal-addr="${escapeHtml(g.addr)}">삭제</button>
                  </td>
                </tr>
              `
                )
                .join("")}
            </tbody>
          </table>
        </div>
      `;

  app.innerHTML = `
  <div style="padding:16px; font-family:system-ui;">
    <h1 style="margin:0 0 12px 0;">Excel Datasheet Generator (ZIP)</h1>

    <div style="display:grid; grid-template-columns: 1.6fr 1fr; gap:16px;">
      <!-- LEFT -->
      <div style="border:1px solid #ddd; border-radius:12px; padding:12px;">
        <h2 style="margin:0 0 8px 0;">1) 템플릿 업로드 & 시트 선택</h2>
        <input id="fileInput" type="file" accept=".xlsx" />

        <div style="margin-top:10px; display:flex; gap:8px; align-items:center; flex-wrap:wrap;">
          <label>Sheet:</label>
          <select id="sheetSelect" ${sheetNames.length ? "" : "disabled"}>
            ${sheetNames.map(n => `<option value="${escapeHtml(n)}" ${n === selectedSheetName ? "selected" : ""}>${escapeHtml(n)}</option>`).join("")}
          </select>
          <button id="loadSheetBtn" ${selectedSheetName ? "" : "disabled"}>시트 로드</button>
        </div>

        <!-- ✅ 표시 범위 컨트롤 -->
        <div style="margin-top:10px; display:flex; gap:8px; align-items:center; flex-wrap:wrap; font-size:12px;">
          <span style="color:#444;">표시:</span>
          <label>행 <input id="maxRInput" type="number" min="10" max="2000" value="${previewMaxR}" style="width:90px;"></label>
          <label>열 <input id="maxCInput" type="number" min="5" max="500" value="${previewMaxC}" style="width:90px;"></label>
          <button id="applyPreviewRangeBtn">적용</button>
          <button id="fitUsedRangeBtn" ${currentSheetRange.rows ? "" : "disabled"} title="ws['!ref'] 사용영역 전체로 맞춤">
            사용영역 전체 (${currentSheetRange.rows || 0}x${currentSheetRange.cols || 0})
          </button>
          <span style="color:#666;">오른쪽/아래 안 보이면 행/열을 늘리세요.</span>
        </div>

        <div id="gridWrap" style="margin-top:10px; max-height:560px; overflow:auto; border:1px solid #eee; border-radius:10px;">
          ${renderGrid(previewGrid, previewMerges, previewColWidths, previewRowHeights)}
        </div>
      </div>

      <!-- RIGHT -->
      <div style="border:1px solid #ddd; border-radius:12px; padding:12px;">
        <h2 style="margin:0 0 8px 0;">2) 셀 편집 등록</h2>

        <div style="padding:10px; border:1px solid #eee; border-radius:10px; background:#fafafa;">
          <div style="display:flex; gap:10px; flex-wrap:wrap;">
            <div><b>시트:</b> ${escapeHtml(selectedSheetName || "-")}</div>
            <div><b>셀:</b> ${escapeHtml(selectedCellAddr || "-")}</div>
          </div>

          <div style="margin-top:10px;">
            <div style="font-weight:700; margin-bottom:6px;">GLOBAL</div>
            <div style="display:flex; gap:6px; flex-wrap:wrap; align-items:center;">
              <input id="globalValueInput" placeholder="공통 문구 입력" style="flex:1; min-width:200px;" />
              <button id="applyGlobalBtn" ${selectedSheetName && selectedCellAddr ? "" : "disabled"}>GLOBAL 등록</button>
              <button id="removeGlobalBtn" ${selectedSheetName && selectedCellAddr ? "" : "disabled"}>GLOBAL 해제</button>
            </div>
          </div>

          <div style="margin-top:12px;">
            <div style="font-weight:700; margin-bottom:6px;">VARIABLE</div>
            <div style="display:flex; gap:6px; flex-wrap:wrap; align-items:center;">
              <input id="fieldKeyInput" placeholder="변수 필드명 (예: ItemNo)" style="flex:1; min-width:200px;" />
              <button id="applyVariableBtn" ${selectedSheetName && selectedCellAddr ? "" : "disabled"}>VARIABLE 등록</button>
              <button id="removeVariableBtn" ${selectedSheetName && selectedCellAddr ? "" : "disabled"}>VARIABLE 해제</button>
            </div>
          </div>
        </div>

        <div style="margin-top:12px;">
          <h3 style="margin:0 0 6px 0;">GLOBAL 요약</h3>
          <div style="display:flex; gap:8px; flex-wrap:wrap; align-items:center;">
            <div style="color:#333;">등록 개수: <b>${globalCount}</b></div>
            <button id="clearAllGlobalBtn">GLOBAL 전체 초기화</button>
          </div>
          ${globalListHtml}
        </div>

        <hr style="margin:12px 0;" />
        <h2 style="margin:0 0 8px 0;">3) VARIABLE 값 입력</h2>
        ${renderVariableFields()}

        <hr style="margin:12px 0;" />
        <h2 style="margin:0 0 8px 0;">4) 파일명 규칙 & Generate</h2>

        <div style="border:1px solid #eee; border-radius:12px; padding:10px;">
          <div style="display:flex; gap:6px; flex-wrap:wrap; align-items:center;">
            <label>Prefix:</label>
            <input id="prefixInput" value="${escapeHtml(fileNamePrefix)}" style="flex:1; min-width:160px;" />
          </div>

          <div style="margin-top:8px; display:flex; gap:6px; flex-wrap:wrap; align-items:center;">
            <label>파일명 필드:</label>
            <select id="fileNameFieldSelect">
              ${
                Object.keys(variableMappings).length
                  ? Object.keys(variableMappings).map(k => `<option value="${escapeHtml(k)}" ${k === fileNameField ? "selected" : ""}>${escapeHtml(k)}</option>`).join("")
                  : `<option value="${escapeHtml(fileNameField)}">${escapeHtml(fileNameField)}</option>`
              }
            </select>
            <span style="color:#666; font-size:12px;">(이 필드 값 개수 = 생성 파일 개수)</span>
          </div>

          <div style="margin-top:8px; display:flex; gap:6px; flex-wrap:wrap; align-items:center;">
            <label>Suffix:</label>
            <input id="suffixInput" value="${escapeHtml(fileNameSuffix)}" style="flex:1; min-width:160px;" />
          </div>

          <div style="margin-top:10px; color:#333;">생성 파일 개수: <b>${fileCount}</b></div>

          <details style="margin-top:8px;">
            <summary style="cursor:pointer;">파일명 미리보기(최대 10개)</summary>
            <pre style="background:#fff; border:1px solid #eee; padding:8px; border-radius:10px; max-height:160px; overflow:auto;">${escapeHtml(filePreview)}</pre>
          </details>

          <div style="margin-top:10px; display:flex; gap:8px; flex-wrap:wrap;">
            <button id="generateZipBtn">Generate ZIP</button>
            <button id="resetAllBtn">전체 초기화(저장 포함)</button>
          </div>
        </div>
      </div>
    </div>
  </div>
  `;

  // ===== events =====
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

  // ✅ 표시 범위 적용
  document.querySelector("#applyPreviewRangeBtn")?.addEventListener("click", () => {
    const r = Number(document.querySelector("#maxRInput")?.value || previewMaxR);
    const c = Number(document.querySelector("#maxCInput")?.value || previewMaxC);
    previewMaxR = Math.max(10, Math.min(2000, r));
    previewMaxC = Math.max(5, Math.min(500, c));
    saveState();
    render();
  });

  // ✅ 사용영역 전체로 맞춤
  document.querySelector("#fitUsedRangeBtn")?.addEventListener("click", () => {
    if (!currentSheetRange.rows || !currentSheetRange.cols) return;
    previewMaxR = Math.min(2000, currentSheetRange.rows);
    previewMaxC = Math.min(500, currentSheetRange.cols);
    saveState();
    render();
  });

  // ✅ 셀 클릭(이벤트 위임: 렌더링 바뀌어도 안정적으로 클릭 잡힘)
  document.querySelector("#gridWrap")?.addEventListener("click", (ev) => {
    const td = ev.target?.closest?.("td[data-cell]");
    if (!td) return;
    selectedCellAddr = td.getAttribute("data-cell") || "";
    render();
  });

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

  document.querySelector("#applyVariableBtn")?.addEventListener("click", onApplyVariable);
  document.querySelector("#removeVariableBtn")?.addEventListener("click", onRemoveVariable);

  document.querySelector("#clearAllGlobalBtn")?.addEventListener("click", () => {
    globalEdits = {};
    saveState();
    render();
  });

  // variable field events
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

  document.querySelectorAll("[data-clearvals]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.getAttribute("data-clearvals");
      ensureField(key);
      variableValues[key] = [];
      saveState();
      render();
    });
  });

  document.querySelectorAll("[data-clearfield]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const key = btn.getAttribute("data-clearfield");
      delete variableMappings[key];
      delete variableValues[key];
      if (fileNameField === key) fileNameField = "ItemNo";
      saveState();
      render();
    });
  });

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

  document.querySelector("#generateZipBtn")?.addEventListener("click", onGenerateZip);

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

  // ✅ 빈칸 포함 grid 생성 + 사용영역 저장
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

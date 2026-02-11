(() => {
  // =========================
  // SPX Toolkit — Sidebar (local)
  // - Bloqueia fora do domínio SPX
  // - Dentro do SPX: tenta pegar Station automaticamente
  // - Dentro do SPX: tenta pegar CSRF automaticamente (meta/cookie)
  // - Remove dependência de "headersJson" (mantém no cfg só por compat)
  // =========================

  const LS_KEY = "spx_toolkit_local_v1";

  const DEFAULTS = {
    stationName: "LM Hub_SC_Chapecó_Eldorado",
    headersJson: "{}", // compat (não usado por padrão)
    lastAction: "ULTIMO_STATUS",
    useCsrfAuto: true,
  };

  // ===== Guard: só roda dentro do SPX =====
  const ALLOWED_HOST_RE = /(^|\.)spx\.shopee\.com\.br$/i;
  const isInSpx = ALLOWED_HOST_RE.test(String(location.hostname || ""));
  if (!isInSpx) {
    const msg =
      "SPX Toolkit: este painel só funciona dentro do sistema SPX (spx.shopee.com.br).\n\nAbra o SPX, faça login, e execute novamente.";
    try {
      alert(msg);
    } catch {}
    console.warn(msg);
    return;
  }

  const RX = {
    BR: /^BR[0-9A-Z]{13}$/i,
    SPX: /^SPX[0-9A-Z]+$/i,
    TO: /^TO[0-9A-Z]+$/i,
    VT: /^VT[0-9A-Z]{13}$/i,
    PS: /^PS[0-9A-Z]+$/i,
    AT: /^AT[0-9A-Z]+$/i,
  };

  const API = {
    trackingInfo: (id) =>
      "https://spx.shopee.com.br/api/fleet_order/order/detail/tracking_info?shipment_id=" +
      encodeURIComponent(id),
    tradeInfo: (id) =>
      "https://spx.shopee.com.br/api/fleet_order/order/detail/trade_info?shipment_id=" +
      encodeURIComponent(id),
    sensitive: (id, field, extra = "") =>
      "https://spx.shopee.com.br/api/fleet_order/order/detail/show_sensitive_data?shipment_id=" +
      encodeURIComponent(id) +
      "&data_field=" +
      encodeURIComponent(field) +
      (extra ? "&" + extra : ""),
    trackListSearch: () =>
      "https://spx.shopee.com.br/api/fleet_order/order/tracking_list/search",
    toOutboundOrderSearch: (to) =>
      "https://spx.shopee.com.br/api/in-station/general_to/outbound/order/search?pageno=1&count=1000000&to_number=" +
      encodeURIComponent(to),
    assignmentDetail: (at) =>
      "https://spx.shopee.com.br/spx_delivery/admin/assignment/assignment_task/detail?assignment_task_id=" +
      encodeURIComponent(at),
    auditTargetListByTask: (vt) =>
      "https://spx.shopee.com.br/api/in-station/lmhub/audit/target/list?page_no=1&count=9999&task_id=" +
      encodeURIComponent(vt),
    auditTargetListByShipment: (vt, br) =>
      "https://spx.shopee.com.br/api/in-station/lmhub/audit/target/list?shipment_id=" +
      encodeURIComponent(br) +
      "&task_id=" +
      encodeURIComponent(vt) +
      "&page_no=1&count=24",
    auditTargetView: (vt, targetId) =>
      "https://spx.shopee.com.br/api/in-station/lmhub/audit/target/view?validation_task_id=" +
      encodeURIComponent(vt) +
      "&target_id=" +
      encodeURIComponent(targetId) +
      "&audit_target_type=2",
    auditParcelList: (vt, targetId, type, page, perPage, auditTargetType = 2) => {
      if (type === "missing") {
        return (
          "https://spx.shopee.com.br/api/in-station/lmhub/audit/parcel/list" +
          "?validation_task_id=" +
          encodeURIComponent(vt) +
          "&target_id=" +
          encodeURIComponent(targetId) +
          "&audit_target_type=" +
          encodeURIComponent(auditTargetType) +
          "&page_no=" +
          page +
          "&count=" +
          perPage +
          "&result=5&shipment_id="
        );
      }
      return (
        "https://spx.shopee.com.br/api/in-station/lmhub/audit/parcel/list" +
        "?validation_task_id=" +
        encodeURIComponent(vt) +
        "&target_id=" +
        encodeURIComponent(targetId) +
        "&audit_target_type=" +
        encodeURIComponent(auditTargetType) +
        "&page_no=" +
        page +
        "&count=" +
        perPage +
        "&parcel_scan_status=2"
      );
    },
  };

  const ACTIONS = [
    { id: "ULTIMO_STATUS", label: "Último status (BR)" },
    { id: "ULTIMA_AT", label: "Última AT (BR)" },
    { id: "MOTIVO_ONHOLD", label: "Motivo do último on-hold (BR)" },
    { id: "ULTIMA_ESTACAO", label: "Última estação (BR)" },
    { id: "HIST_ESTACOES", label: "Histórico de estações (BR)" },
    { id: "AGING", label: "Aging na station (BR)" },
    { id: "NOME_ITEM", label: "Nome do item (SKU/name) (BR)" },
    { id: "TO_PUXAR_BRS", label: "Listar BRs por TO" },
    { id: "RETURNS", label: "Return's (SPX/BR → TO/LH)" },
    { id: "AUDIT_MM", label: "Conferência Missing/Missort (VT)" },
  ];

  // =========================
  // Config
  // =========================
  function loadCfg() {
    try {
      const raw = localStorage.getItem(LS_KEY);
      const obj = raw ? JSON.parse(raw) : {};
      return { ...DEFAULTS, ...(obj || {}) };
    } catch {
      return { ...DEFAULTS };
    }
  }

  function saveCfg(patch) {
    const next = { ...loadCfg(), ...(patch || {}) };
    localStorage.setItem(LS_KEY, JSON.stringify(next));
    return next;
  }

  // =========================
  // Auto-detect helpers (SPX)
  // =========================
  function detectStationFromDOM() {
    const cand = [];

    // 1) texto do topo / sidebar (ex: "LM Hub_SC_...")
    const textNodes = [
      ...document.querySelectorAll("body *"),
    ];
    // varredura leve: só elementos comuns que tendem a conter station
    const picks = [];
    for (const el of textNodes) {
      if (!el) continue;
      const tag = el.tagName;
      if (!tag) continue;
      if (tag === "SCRIPT" || tag === "STYLE" || tag === "NOSCRIPT") continue;
      if (tag === "INPUT" || tag === "TEXTAREA") continue;
      const s = (el.textContent || "").trim();
      if (!s) continue;
      if (s.length > 120) continue;
      if (/^(LM|FM)\s*Hub[_\s]/i.test(s) || /^[A-Z]{1,3}\s*Hub[_\s]/.test(s) || /Hub_SC_/i.test(s)) {
        picks.push(s);
      }
      if (picks.length >= 30) break;
    }
    cand.push(...picks);

    // Normaliza pra formato provável "LM Hub_SC_..."
    const normalized = cand
      .map((s) => s.replace(/\s+/g, " ").trim())
      .filter(Boolean);

    // tenta achar exatamente com underscore
    const bestUnderscore = normalized.find((s) => /Hub_/.test(s));
    if (bestUnderscore) return bestUnderscore;

    // fallback: o primeiro
    return normalized[0] || "";
  }

  function detectStation() {
    const cfg = loadCfg();
    const dom = detectStationFromDOM();
    const st = dom || cfg.stationName || DEFAULTS.stationName;
    if (st && st !== cfg.stationName) saveCfg({ stationName: st });
    return st;
  }

  function getCsrfToken() {
    // meta
    const meta =
      document.querySelector('meta[name="csrf-token"]')?.content ||
      document.querySelector('meta[name="csrfToken"]')?.content ||
      document.querySelector('meta[name="csrf"]')?.content ||
      "";
    if (meta) return String(meta).trim();

    // cookie comum
    const m = document.cookie.match(/(?:^|;\s*)csrftoken=([^;]+)/i);
    if (m && m[1]) return decodeURIComponent(m[1]);

    // outros nomes possíveis
    const m2 = document.cookie.match(/(?:^|;\s*)csrf_token=([^;]+)/i);
    if (m2 && m2[1]) return decodeURIComponent(m2[1]);

    return "";
  }

  function stationName() {
    return detectStation() || DEFAULTS.stationName;
  }

  // Mantém por compat, mas não usa por padrão
  function safeHeadersFromCfg() {
    const cfg = loadCfg();
    let h = {};
    try {
      h = JSON.parse(cfg.headersJson || "{}") || {};
    } catch {
      h = {};
    }

    const blocked = new Set([
      "cookie",
      "host",
      "origin",
      "referer",
      "user-agent",
      "content-length",
      "connection",
      "accept-encoding",
    ]);

    const out = {};
    for (const [k, v] of Object.entries(h)) {
      const kk = String(k || "").toLowerCase().trim();
      if (!kk || blocked.has(kk)) continue;
      if (v == null) continue;
      out[k] = String(v);
    }
    return out;
  }

  function uniq(arr) {
    const s = new Set();
    const out = [];
    for (const x of arr) {
      if (!s.has(x)) {
        s.add(x);
        out.push(x);
      }
    }
    return out;
  }

  function flatTracking(list) {
    return (list || []).reduce((acc, n) => {
      acc.push(n);
      if (Array.isArray(n.children)) acc.push(...flatTracking(n.children));
      return acc;
    }, []);
  }

  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  function fmtEpoch(sec) {
    const n = Number(sec);
    if (!n || !isFinite(n)) return "";
    const d = new Date(n * 1000);
    return `${pad2(d.getDate())}/${pad2(d.getMonth() + 1)}/${d.getFullYear()} ${pad2(
      d.getHours()
    )}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
  }

  // =========================
  // HTTP (auto dentro do SPX)
  // =========================
  async function apiFetch(url, opt = {}) {
    const cfg = loadCfg();

    // base headers: mínimo
    const headers = {};

    // se for POST json, garante content-type
    if ((opt.method || "GET").toUpperCase() === "POST") {
      if (!opt.headers || !Object.keys(opt.headers).some((k) => String(k).toLowerCase() === "content-type")) {
        headers["content-type"] = "application/json;charset=UTF-8";
      }
    }

    // CSRF auto (se existir)
    if (cfg.useCsrfAuto) {
      const csrf = getCsrfToken();
      if (csrf) {
        // nomes comuns: x-csrftoken / x-csrf-token
        headers["x-csrftoken"] = csrf;
        headers["x-csrf-token"] = csrf;
      }
    }

    // compat: se você quiser insistir em algum header extra via cfg
    const cfgHeaders = safeHeadersFromCfg();
    for (const [k, v] of Object.entries(cfgHeaders)) headers[k] = v;

    // headers do opt têm prioridade
    for (const [k, v] of Object.entries(opt.headers || {})) headers[k] = v;

    const res = await fetch(url, {
      method: opt.method || "GET",
      credentials: "include",
      headers,
      body: opt.body,
    });

    const text = await res.text();

    let json = null;
    try {
      json = JSON.parse(text);
    } catch {
      json = { ok: false, status: res.status, error: "INVALID_JSON", raw: text };
    }

    if (!res.ok) {
      const msg =
        (json && (json.message || json.msg || json.error)) || `HTTP ${res.status} em ${url}`;
      const err = new Error(msg);
      err.status = res.status;
      err.url = url;
      err.payload = json;
      throw err;
    }

    return json;
  }

  async function fetchPool(jobs, { concurrency = 6, onProgress } = {}) {
    const out = new Array(jobs.length);
    let idx = 0;
    let done = 0;
    const workers = Math.max(1, Math.min(concurrency, jobs.length || 1));

    await Promise.all(
      new Array(workers).fill(0).map(async () => {
        while (true) {
          const i = idx++;
          if (i >= jobs.length) return;
          try {
            const j = jobs[i];
            out[i] = await apiFetch(j.url, j.opt || {});
          } catch (e) {
            out[i] = { ok: false, _err: true, error: String(e.message || e) };
          } finally {
            done++;
            if (onProgress) onProgress(done, jobs.length, i);
          }
        }
      })
    );

    return out;
  }

  function parseCodes(raw) {
    return String(raw || "")
      .split(/[\n,;\t\r ]+/g)
      .map((s) => s.trim().toUpperCase())
      .filter(Boolean);
  }

  function escapeHtml(s) {
    return String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  }

  // =========================
  // Sidebar UI (injeta tudo)
  // =========================
  (function mountUI() {
    const existing = document.getElementById("spx-toolkit-panel");
    if (existing) existing.remove();

    const style = document.createElement("style");
    style.textContent = `
#spx-toolkit-panel{position:fixed;right:0;top:0;height:100vh;z-index:2147483647;width:440px;max-width:calc(100vw - 24px);font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif}
#spx-toolkit-panel .wrap{height:100%;background:#0b1220;color:#e5e7eb;border-left:1px solid rgba(255,255,255,.12);box-shadow:-18px 0 60px rgba(0,0,0,.45);display:flex;flex-direction:column}
#spx-toolkit-panel .hdr{display:flex;align-items:center;justify-content:space-between;padding:12px;border-bottom:1px solid rgba(255,255,255,.10)}
#spx-toolkit-panel .hdr b{font-size:13px;letter-spacing:.2px}
#spx-toolkit-panel .hdr .btns{display:flex;gap:8px}
#spx-toolkit-panel button{cursor:pointer;border:1px solid rgba(255,255,255,.14);background:rgba(255,255,255,.08);color:#e5e7eb;border-radius:10px;padding:8px 10px;font-size:12px}
#spx-toolkit-panel button:hover{background:rgba(255,255,255,.12)}
#spx-toolkit-panel .body{padding:12px;display:flex;flex-direction:column;gap:10px;overflow:auto}
#spx-toolkit-panel textarea{width:100%;min-height:110px;resize:vertical;background:#07101f;color:#e5e7eb;border:1px solid rgba(255,255,255,.14);border-radius:10px;padding:10px;font-size:12px;line-height:1.35}
#spx-toolkit-panel select,#spx-toolkit-panel input{width:100%;background:#07101f;color:#e5e7eb;border:1px solid rgba(255,255,255,.14);border-radius:10px;padding:9px 10px;font-size:12px}
#spx-toolkit-panel .row{display:flex;gap:10px}
#spx-toolkit-panel .row > *{flex:1}
#spx-toolkit-panel .log{background:#07101f;border:1px solid rgba(255,255,255,.14);border-radius:10px;padding:10px;font-size:12px;max-height:130px;overflow:auto;white-space:pre-wrap}
#spx-toolkit-panel .tblwrap{border:1px solid rgba(255,255,255,.14);border-radius:10px;overflow:auto;max-height:360px}
#spx-toolkit-panel table{border-collapse:collapse;width:100%;font-size:12px}
#spx-toolkit-panel th,#spx-toolkit-panel td{border-bottom:1px solid rgba(255,255,255,.08);padding:8px 10px;vertical-align:top}
#spx-toolkit-panel th{position:sticky;top:0;background:#0b1220;text-align:left}
#spx-toolkit-panel .muted{opacity:.85;font-size:11px}
#spx-toolkit-panel .meta{display:flex;justify-content:space-between;align-items:center;font-size:12px;opacity:.95}
#spx-toolkit-panel .pill{display:inline-flex;gap:8px;align-items:center;background:rgba(255,255,255,.07);border:1px solid rgba(255,255,255,.12);border-radius:999px;padding:6px 10px;font-size:11px}
#spx-toolkit-panel .danger{color:#fca5a5}
    `;
    document.documentElement.appendChild(style);

    const root = document.createElement("div");
    root.id = "spx-toolkit-panel";
    root.innerHTML = `
      <div class="wrap">
        <div class="hdr">
          <div>
            <b>SPX Toolkit</b>
            <div class="muted">Sidebar • Local • auto-auth</div>
          </div>
          <div class="btns">
            <button data-act="cfg">Config</button>
            <button data-act="min">Min</button>
            <button data-act="x">Fechar</button>
          </div>
        </div>

        <div class="body" data-body="1">
          <div class="meta">
            <div class="pill">Station: <span data-el="station"></span></div>
            <div class="pill">CSRF: <span data-el="csrf"></span></div>
          </div>

          <div class="row">
            <select data-el="action"></select>
            <button data-el="run">Rodar</button>
          </div>

          <textarea data-el="input" placeholder="Cole BRs / SPX / TO / VT (um por linha)"></textarea>

          <div class="row">
            <button data-el="copy">Copiar CSV</button>
            <button data-el="dl">Baixar CSV</button>
          </div>

          <div class="log" data-el="log"></div>

          <div class="meta">
            <div>Resultado</div>
            <div data-el="rows">Linhas: 0</div>
          </div>

          <div class="tblwrap" data-el="tblwrap"></div>

          <div class="muted">
            Dica: você precisa estar logado no SPX. O browser envia cookie automaticamente.
          </div>
        </div>
      </div>
    `;
    document.documentElement.appendChild(root);

    // Elements
    const elAction = root.querySelector('[data-el="action"]');
    const btnRun = root.querySelector('[data-el="run"]');
    const input = root.querySelector('[data-el="input"]');
    const log = root.querySelector('[data-el="log"]');
    const tableWrap = root.querySelector('[data-el="tblwrap"]');
    const elRows = root.querySelector('[data-el="rows"]');
    const btnCopy = root.querySelector('[data-el="copy"]');
    const btnDownload = root.querySelector('[data-el="dl"]');
    const elStation = root.querySelector('[data-el="station"]');
    const elCsrf = root.querySelector('[data-el="csrf"]');

    // Close/min
    root.querySelector('[data-act="x"]').onclick = () => root.remove();
    root.querySelector('[data-act="min"]').onclick = () => {
      const body = root.querySelector('[data-body="1"]');
      body.style.display = body.style.display === "none" ? "" : "none";
    };

    // Config modal (simples, embutido)
    root.querySelector('[data-act="cfg"]').onclick = () => openConfig();

    let lastCSV = "";

    function logLine(msg) {
      const d = new Date();
      const ts = `${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`;
      log.textContent = `[${ts}] ${msg}\n` + log.textContent;
    }

    function renderTable(header, rows) {
      const th = header.map((h) => `<th>${escapeHtml(h)}</th>`).join("");
      const body = rows
        .slice(0, 3000)
        .map((r) => `<tr>${r.map((c) => `<td>${escapeHtml(c)}</td>`).join("")}</tr>`)
        .join("");
      tableWrap.innerHTML = `<table><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table>`;
    }

    function toCSV(header, rows) {
      const esc = (v) => {
        const s = String(v ?? "");
        return /[,"\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
      };
      return [header.map(esc).join(","), ...rows.map((r) => r.map(esc).join(","))].join("\n");
    }

    function downloadText(filename, text) {
      const blob = new Blob([text], { type: "text/csv;charset=utf-8" });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 1500);
    }

    btnCopy.onclick = async () => {
      if (!lastCSV) return logLine("Nada para copiar ainda.");
      try {
        await navigator.clipboard.writeText(lastCSV);
        logLine("CSV copiado.");
      } catch (e) {
        logLine("Falha ao copiar: " + String(e.message || e));
      }
    };

    btnDownload.onclick = () => {
      if (!lastCSV) return logLine("Nada para baixar ainda.");
      downloadText("spx_toolkit.csv", lastCSV);
      logLine("Download iniciado.");
    };

    function refreshBadges() {
      const st = stationName();
      const csrf = getCsrfToken();
      elStation.textContent = st || "-";
      elCsrf.textContent = csrf ? "OK" : "—";
    }

    // Config UI
    function openConfig() {
      const old = document.getElementById("spx-toolkit-cfg");
      if (old) old.remove();

      const cfg = loadCfg();

      const overlay = document.createElement("div");
      overlay.id = "spx-toolkit-cfg";
      overlay.style.cssText =
        "position:fixed;inset:0;z-index:2147483647;background:rgba(0,0,0,.55);display:flex;align-items:center;justify-content:center;padding:18px";

      const box = document.createElement("div");
      box.style.cssText =
        "width:520px;max-width:calc(100vw - 36px);background:#0b1220;color:#e5e7eb;border:1px solid rgba(255,255,255,.12);border-radius:14px;box-shadow:0 20px 60px rgba(0,0,0,.45);overflow:hidden";

      box.innerHTML = `
        <div style="padding:12px;border-bottom:1px solid rgba(255,255,255,.10);display:flex;justify-content:space-between;align-items:center">
          <b>Config</b>
          <button data-x="1">Fechar</button>
        </div>
        <div style="padding:12px;display:flex;flex-direction:column;gap:10px">
          <label style="font-size:12px">Station name (auto detect liga sozinho; você pode ajustar)</label>
          <input data-st="1" value="${escapeHtml(cfg.stationName || DEFAULTS.stationName)}" />
          <label style="font-size:12px">CSRF automático</label>
          <select data-csrf="1">
            <option value="1">Ativo</option>
            <option value="0">Desligado</option>
          </select>

          <label style="font-size:12px">Headers JSON (opcional/compat; normalmente não precisa)</label>
          <textarea data-h="1" style="min-height:140px">${escapeHtml(cfg.headersJson || "{}")}</textarea>

          <div style="display:flex;gap:10px;justify-content:flex-end">
            <button data-save="1">Salvar</button>
          </div>

          <div style="font-size:11px;opacity:.85">
            Se der 401 / login expired: faça login de novo no SPX.
          </div>
        </div>
      `;
      overlay.appendChild(box);
      document.documentElement.appendChild(overlay);

      const close = () => overlay.remove();
      box.querySelector('[data-x="1"]').onclick = close;
      overlay.onclick = (e) => {
        if (e.target === overlay) close();
      };

      const selCsrf = box.querySelector('[data-csrf="1"]');
      selCsrf.value = cfg.useCsrfAuto ? "1" : "0";

      box.querySelector('[data-save="1"]').onclick = () => {
        const st = box.querySelector('[data-st="1"]').value.trim();
        const hj = (box.querySelector('[data-h="1"]').value || "").trim() || "{}";
        const useCsrfAuto = selCsrf.value === "1";
        saveCfg({ stationName: st || DEFAULTS.stationName, headersJson: hj, useCsrfAuto });
        logLine("Config salva no localStorage.");
        refreshBadges();
        close();
      };
    }

    // =========================
    // Audit cache: rota correta
    // =========================
    const shipmentTargetCache = {};
    async function getCorrectRoute(vt, br, currentAT) {
      const key = `${vt}||${br}`;
      if (!shipmentTargetCache[key]) {
        try {
          const j = await apiFetch(API.auditTargetListByShipment(vt, br));
          shipmentTargetCache[key] = j?.data?.list || [];
        } catch {
          shipmentTargetCache[key] = [];
        }
      }
      const list = shipmentTargetCache[key];
      const other = list.find((t) => t.target_id !== currentAT);
      if (other?.binding_entity) return other.binding_entity;
      const self = list.find((t) => t.target_id === currentAT);
      return self?.binding_entity || "";
    }

    // =========================
    // Ações
    // =========================
    async function act_TO_PUXAR_BRS(codes) {
      const tos = uniq(codes.filter((c) => RX.TO.test(c)));
      if (!tos.length) throw new Error("Nenhum TO válido.");

      logLine(`TO(s): ${tos.length}`);
      const jobs = tos.map((to) => ({ url: API.toOutboundOrderSearch(to) }));

      const resps = await fetchPool(jobs, {
        concurrency: 4,
        onProgress: (d, t) => logLine(`TO: ${d}/${t}`),
      });

      const header = [
        "TO",
        "BR",
        "station_name",
        "third_party_sorting_code",
        "status",
        "receiver_name",
        "ctime",
        "mtime",
      ];
      const rows = [];

      for (let i = 0; i < tos.length; i++) {
        const to = tos[i];
        const r = resps[i];
        const list = Array.isArray(r?.data?.list) ? r.data.list : [];
        if (!list.length) {
          rows.push([to, "", "", "", "", "", "", ""]);
          continue;
        }
        for (const it of list) {
          const br = it?.sls_tracking_number || it?.shipment_id || it?.fleet_order_id || "";
          rows.push([
            to,
            String(br),
            it?.station_name || "",
            it?.third_party_sorting_code || "",
            typeof it?.status === "number" ? it.status : it?.status || "",
            it?.receiver_name || "",
            fmtEpoch(it?.ctime),
            fmtEpoch(it?.mtime),
          ]);
        }
      }

      return { header, rows };
    }

    async function act_RETURNS(codes) {
      const station = stationName();
      const ids = uniq(codes.filter((c) => RX.BR.test(c) || RX.SPX.test(c))).slice(0, 200);
      if (!ids.length) throw new Error("Nenhum BR/SPX válido.");

      logLine(`Return's: ${ids.length} | station=${station}`);
      const jobs = ids.map((id) => ({ url: API.trackingInfo(id) }));

      const resps = await fetchPool(jobs, {
        concurrency: 6,
        onProgress: (d, t) => (d % 10 === 0 ? logLine(`Return's: ${d}/${t}`) : null),
      });

      const RX_MSG =
        /(?:Parcel\s+\[(TO[0-9A-Z]+)\]\s+added\s+into\s+LH\s+Task\s+\[([^\]]+)\])|(?:Parcel's\s+TO\s+\[(TO[0-9A-Z]+)\]\s+adding\s+into\s+LH\s+Task\s+\[([^\]]+)\])/i;

      const header = ["SPX TN", "DATA", "TO", "LH"];
      const rows = ids.map((id, i) => {
        const all = flatTracking(resps[i]?.data?.tracking_list || []);
        if (!all.length) return [id, "", "", ""];
        const cand = all.filter(
          (n) => String(n.station_name || "") === station && RX_MSG.test(String(n.message || ""))
        );
        if (!cand.length) return [id, "", "", ""];
        const last = cand.reduce((m, v) => (Number(v.timestamp) > Number(m.timestamp) ? v : m), cand[0]);
        const m = String(last.message || "").match(RX_MSG);
        const to = (m && (m[1] || m[3])) || "";
        const lh = (m && (m[2] || m[4])) || "";
        return [id, Number(last.timestamp) ? fmtEpoch(Number(last.timestamp)) : "", to, lh];
      });

      return { header, rows };
    }

    async function act_AUDIT_MM(codes) {
      const vts = uniq(codes.filter((c) => RX.VT.test(c))).slice(0, 2);
      if (!vts.length) throw new Error("Cole 1 ou 2 VT(s).");

      logLine(`Audit: VTs=${vts.join(", ")}`);

      const brRows = [];
      const atSet = new Set();
      const targetToVT = {};

      for (let vi = 0; vi < vts.length; vi++) {
        const vt = vts[vi];
        logLine(`VT ${vt}: buscando targets...`);
        const tlist = await apiFetch(API.auditTargetListByTask(vt));
        const targets = (tlist?.data?.list || []).filter(
          (t) => (t.missing_qty && t.missing_qty > 0) || (t.missort_qty && t.missort_qty > 0)
        );

        logLine(`VT ${vt}: targets=${targets.length}`);

        for (const target of targets) {
          const targetId = target.target_id;
          const auditTargetType = target.audit_target_type || 2;
          atSet.add(targetId);
          targetToVT[targetId] = vt;

          const perPage = 200;

          if (target.missing_qty && target.missing_qty > 0) {
            logLine(`AT ${targetId}: missing...`);
            let page = 1;
            while (true) {
              const r = await apiFetch(
                API.auditParcelList(vt, targetId, "missing", page, perPage, auditTargetType)
              );
              const total = r?.data?.total || 0;
              const list = r?.data?.list || [];
              for (const it of list) brRows.push({ vt, targetId, type: "missing", br: it.shipment_id || "" });
              if (!list.length || page * perPage >= total) break;
              page++;
              await new Promise((r) => setTimeout(r, 120));
            }
          }

          if (target.missort_qty && target.missort_qty > 0) {
            logLine(`AT ${targetId}: missort...`);
            let page = 1;
            while (true) {
              const r = await apiFetch(
                API.auditParcelList(vt, targetId, "missort", page, perPage, auditTargetType)
              );
              const total = r?.data?.total || 0;
              const list = r?.data?.list || [];
              const only7 = list.filter((it) => it.validation_status === 7);
              for (const it of only7) brRows.push({ vt, targetId, type: "missort", br: it.shipment_id || "" });
              if (!list.length || page * perPage >= total) break;
              page++;
              await new Promise((r) => setTimeout(r, 120));
            }
          }
        }
      }

      const atArray = Array.from(atSet);
      logLine(`Audit: BRs=${brRows.length} | ATs=${atArray.length}`);

      // assignments
      const assignJobs = atArray.map((at) => ({ url: API.assignmentDetail(at) }));
      const assigns = await fetchPool(assignJobs, {
        concurrency: 4,
        onProgress: (d, t) => (d % 5 === 0 ? logLine(`Assignments: ${d}/${t}`) : null),
      });
      const assignMap = {};
      atArray.forEach((at, i) => (assignMap[at] = assigns[i]?.data || {}));

      // target/view
      const tvJobs = atArray.map((at) => ({ url: API.auditTargetView(targetToVT[at], at) }));
      const tviews = await fetchPool(tvJobs, {
        concurrency: 4,
        onProgress: (d, t) => (d % 5 === 0 ? logLine(`Target/view: ${d}/${t}`) : null),
      });
      const tvMap = {};
      atArray.forEach((at, i) => (tvMap[at] = tviews[i]?.data || {}));

      const header = [
        "DATA_HORA",
        "VT",
        "AT",
        "BR",
        "ROTA_ENCONTRADA",
        "ROTA_CORRETA",
        "OPERADOR",
        "DRIVER_ID",
        "DRIVER_NAME",
        "TIPO_DE_ERRO",
      ];

      const rows = [];
      for (let i = 0; i < brRows.length; i++) {
        if (i % 120 === 0) logLine(`Montando linhas: ${i}/${brRows.length}`);
        const r = brRows[i];
        const asg = assignMap[r.targetId] || {};
        const tv = tvMap[r.targetId] || {};
        const dt = asg.assigned_time ? fmtEpoch(asg.assigned_time) : "";
        const rotaCorreta = await getCorrectRoute(r.vt, r.br, r.targetId);

        rows.push([
          dt,
          r.vt,
          r.targetId,
          r.br,
          tv.binding_entity || "",
          rotaCorreta,
          tv.validation_operator || "",
          asg.driver_id || "",
          asg.driver_name || "",
          r.type,
        ]);
      }

      return { header, rows };
    }

    async function act_BR_simple(action, brs) {
      const ids = uniq(brs.filter((c) => RX.BR.test(c)));
      if (!ids.length) throw new Error("Nenhum BR válido.");

      const station = stationName();

      if (action === "ULTIMO_STATUS") {
        logLine(`TrackingInfo: ${ids.length}`);
        const resps = await fetchPool(ids.map((id) => ({ url: API.trackingInfo(id) })), {
          concurrency: 6,
          onProgress: (d, t) => (d % 10 === 0 ? logLine(`Status: ${d}/${t}`) : null),
        });

        const header = ["BR", "RESULTADO"];
        const rows = ids.map((id, i) => {
          const all = flatTracking(resps[i]?.data?.tracking_list || []);
          if (!all.length) return [id, ""];
          const last = all.reduce((m, v) => (Number(v.timestamp) > Number(m.timestamp) ? v : m), all[0]);
          return [id, last.message || ""];
        });
        return { header, rows };
      }

      if (action === "ULTIMA_AT") {
        logLine(`TrackingInfo: ${ids.length}`);
        const resps = await fetchPool(ids.map((id) => ({ url: API.trackingInfo(id) })), {
          concurrency: 6,
          onProgress: (d, t) => (d % 10 === 0 ? logLine(`AT: ${d}/${t}`) : null),
        });

        const header = ["BR", "RESULTADO"];
        const rows = ids.map((id, i) => {
          const all = flatTracking(resps[i]?.data?.tracking_list || []).filter((n) => {
            const statusStr = String(n.event_type || n.event_code || n.biz_code || n.status_text || "");
            const msg = String(n.message || "");
            const isAssign =
              /LMHub_Assign(?:ed|ing)/i.test(statusStr) || /Pedido em processamento na Assignment Task/i.test(msg);
            const hasAT = /\[(AT[0-9A-Z]+)\]/i.test(msg);
            return isAssign && hasAT;
          });
          if (!all.length) return [id, "❌ não encontrado"];
          const last = all.reduce((m, v) => (Number(v.timestamp) > Number(m.timestamp) ? v : m), all[0]);
          const m = String(last.message || "").match(/\[(AT[0-9A-Z]+)\]/i);
          return [id, (m ? m[1] : "") || "❌ não encontrado"];
        });
        return { header, rows };
      }

      if (action === "MOTIVO_ONHOLD") {
        logLine(`TrackingInfo: ${ids.length}`);
        const resps = await fetchPool(ids.map((id) => ({ url: API.trackingInfo(id) })), {
          concurrency: 6,
          onProgress: (d, t) => (d % 10 === 0 ? logLine(`On-hold: ${d}/${t}`) : null),
        });

        const header = ["BR", "RESULTADO"];
        const rows = ids.map((id, i) => {
          const hold = flatTracking(resps[i]?.data?.tracking_list || []).filter(
            (n) => typeof n.message === "string" && /^\s*Pedido em espera\s*:/i.test(n.message)
          );
          if (!hold.length) return [id, "❌ não encontrado"];
          const last = hold.reduce((m, v) => (Number(v.timestamp) > Number(m.timestamp) ? v : m), hold[0]);
          const matches = String(last.message || "").match(/\[([^\]]+)\]/g);
          if (!matches?.length) return [id, "❌ não encontrado"];
          return [id, matches[matches.length - 1].replace(/^\[|\]$/g, "").trim() || "❌ não encontrado"];
        });
        return { header, rows };
      }

      if (action === "ULTIMA_ESTACAO") {
        logLine(`TrackingListSearch: ${ids.length}`);
        const jobs = ids.map((id) => ({
          url: API.trackListSearch(),
          opt: {
            method: "POST",
            headers: { "content-type": "application/json;charset=UTF-8" },
            body: JSON.stringify({ shipment_id: id, count: 24, page_no: 1 }),
          },
        }));

        const resps = await fetchPool(jobs, {
          concurrency: 6,
          onProgress: (d, t) => (d % 10 === 0 ? logLine(`Estação: ${d}/${t}`) : null),
        });

        const header = ["BR", "RESULTADO"];
        const rows = ids.map((id, i) => {
          const list = resps[i]?.data?.list;
          if (!list?.length) return [id, "❌ sem tracking"];
          return [id, list[0]?.station_name || "❌ sem station_name"];
        });
        return { header, rows };
      }

      if (action === "HIST_ESTACOES") {
        logLine(`TrackingInfo: ${ids.length}`);
        const resps = await fetchPool(ids.map((id) => ({ url: API.trackingInfo(id) })), {
          concurrency: 6,
          onProgress: (d, t) => (d % 10 === 0 ? logLine(`Hist: ${d}/${t}`) : null),
        });

        const header = ["BR", "RESULTADO"];
        const rows = ids.map((id, i) => {
          const all = flatTracking(resps[i]?.data?.tracking_list || []);
          all.sort((a, b) => Number(a.timestamp) - Number(b.timestamp));
          const seq = all
            .map((n) => n.station_name || "")
            .filter((s, k, arr) => s && (k === 0 || s !== arr[k - 1]));
          return [id, seq.join(" > ")];
        });
        return { header, rows };
      }

      if (action === "AGING") {
        logLine(`TrackingInfo: ${ids.length} | station=${station}`);
        const resps = await fetchPool(ids.map((id) => ({ url: API.trackingInfo(id) })), {
          concurrency: 6,
          onProgress: (d, t) => (d % 10 === 0 ? logLine(`Aging: ${d}/${t}`) : null),
        });

        const header = ["BR", "RESULTADO"];
        const rows = ids.map((id, i) => {
          const all = flatTracking(resps[i]?.data?.tracking_list || []);
          if (!all.length) return [id, "❌ sem tracking"];
          const hits = all.filter((n) => String(n.station_name || "") === station && Number(n.timestamp));
          if (!hits.length) return [id, "❌ sem registros na estação"];
          let first = hits[0];
          let last = hits[0];
          for (let k = 1; k < hits.length; k++) {
            const cur = hits[k];
            if (Number(cur.timestamp) < Number(first.timestamp)) first = cur;
            if (Number(cur.timestamp) > Number(last.timestamp)) last = cur;
          }
          const durSec = Number(last.timestamp) - Number(first.timestamp);
          if (!isFinite(durSec) || durSec < 0) return [id, "❌ duração inválida"];
          const h = Math.round((durSec / 3600) * 100) / 100;
          return [id, String(h).replace(".", ",") + " h"];
        });
        return { header, rows };
      }

      if (action === "NOME_ITEM") {
        logLine(`TradeInfo: ${ids.length}`);
        const trade = await fetchPool(ids.map((id) => ({ url: API.tradeInfo(id) })), {
          concurrency: 6,
          onProgress: (d, t) => (d % 10 === 0 ? logLine(`Trade: ${d}/${t}`) : null),
        });

        const skuIds = trade.map((j) => j?.data?.sku_list?.[0]?.id || null);

        const reqIdx = [];
        const jobs = [];
        for (let i = 0; i < ids.length; i++) {
          const sku = skuIds[i];
          if (sku) {
            reqIdx.push(i);
            jobs.push({ url: API.sensitive(ids[i], "name", "id=" + encodeURIComponent(String(sku))) });
          }
        }

        const names = new Array(ids.length).fill("❌ nenhum SKU encontrado");

        if (jobs.length) {
          logLine(`Sensitive(name): ${jobs.length}`);
          const resps = await fetchPool(jobs, {
            concurrency: 6,
            onProgress: (d, t) => (d % 10 === 0 ? logLine(`Name: ${d}/${t}`) : null),
          });

          for (let k = 0; k < resps.length; k++) {
            const i = reqIdx[k];
            const val = resps[k]?.data?.data_detail;
            names[i] = val == null || val === "" ? "❌ resposta inesperada" : String(val);
          }
        }

        const header = ["BR", "RESULTADO"];
        const rows = ids.map((id, i) => [id, names[i]]);
        return { header, rows };
      }

      throw new Error("Ação não implementada.");
    }

    // =========================
    // Wire UI
    // =========================
    function init() {
      // actions
      elAction.innerHTML = "";
      for (const a of ACTIONS) {
        const opt = document.createElement("option");
        opt.value = a.id;
        opt.textContent = a.label;
        elAction.appendChild(opt);
      }
      const cfg = loadCfg();
      elAction.value = cfg.lastAction || DEFAULTS.lastAction;

      refreshBadges();
      logLine("Sidebar carregada.");
    }

    btnRun.onclick = async () => {
      const action = elAction.value;
      saveCfg({ lastAction: action });

      const codes = parseCodes(input.value);
      if (!codes.length) return logLine("Cole pelo menos 1 código.");

      btnRun.disabled = true;
      btnRun.textContent = "Rodando...";
      elRows.textContent = "Linhas: ...";

      try {
        refreshBadges();
        logLine("Ação: " + action);

        let result;
        if (action === "TO_PUXAR_BRS") result = await act_TO_PUXAR_BRS(codes);
        else if (action === "RETURNS") result = await act_RETURNS(codes);
        else if (action === "AUDIT_MM") result = await act_AUDIT_MM(codes);
        else result = await act_BR_simple(action, codes);

        renderTable(result.header, result.rows);
        lastCSV = toCSV(result.header, result.rows);
        elRows.textContent = `Linhas: ${result.rows.length}`;
        logLine("OK: linhas=" + result.rows.length);
      } catch (e) {
        logLine("ERRO: " + String(e.message || e));
        elRows.innerHTML = `Linhas: <span class="danger">erro</span>`;
      } finally {
        btnRun.disabled = false;
        btnRun.textContent = "Rodar";
      }
    };

    init();
  })();
})();

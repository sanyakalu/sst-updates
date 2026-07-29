/* SST Update Generator — glue between the UI and Pyodide. */

const PYODIDE_INDEX = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";
const WORK    = "/work";
const OUT_DIR = WORK + "/output";

// ── DOM refs ──────────────────────────────────────────────────────────────────

const els = {
  prevFile:    document.getElementById("prevFile"),
  updFile:     document.getElementById("updFile"),
  chosenPrev:  document.getElementById("chosen-prev"),
  chosenUpd:   document.getElementById("chosen-upd"),
  monthGrid:   document.getElementById("month-grid"),
  year:        document.getElementById("year"),
  userEmail:     document.getElementById("userEmail"),
  emailErr:      document.getElementById("emailErr"),
  emailReveal:   document.getElementById("emailReveal"),
  createSticr:   document.getElementById("createSticr"),
  createQualReg: document.getElementById("createQualReg"),
  run:           document.getElementById("run"),
  status:        document.getElementById("status"),
  statusText:    document.getElementById("statusText"),
  spin:          document.getElementById("spin"),
  download:      document.getElementById("download"),
  downloadCsv:   document.getElementById("download-csv"),
  log:         document.getElementById("log"),
  verBanner:   document.getElementById("ver-banner"),
  verBody:     document.getElementById("ver-body"),
  verDismiss:  document.getElementById("ver-dismiss"),
};

// ── State ─────────────────────────────────────────────────────────────────────

let pyodide       = null;
let pyReady       = false;
let selectedMonth = null;
let updBytes      = null;
let updFileName   = "update.txt";
let verBlocking   = false;  // true when version check found issues

// ── STICR template fields (loaded from repo JSON at boot) ────────────────────
let templateFields = [];

// Fields set explicitly in createSticrs — must not be duplicated from the JSON
const STICR_EXCLUDED = new Set([
  "System.Title", "System.Description", "System.AreaPath",
  "System.TeamProject", "System.IterationPath", "System.AssignedTo",
  "System.Id", "System.Rev", "System.WorkItemType",
  "System.CreatedDate", "System.CreatedBy", "System.ChangedDate",
  "System.ChangedBy", "System.BoardColumn", "System.BoardColumnDone",
  "System.CommentCount", "System.AuthorizedDate", "System.Watermark",
  "Microsoft.VSTS.Common.StateChangeDate",
  "Microsoft.VSTS.Common.ResolvedDate", "Microsoft.VSTS.Common.ResolvedBy",
  "Microsoft.VSTS.Common.ClosedDate", "Microsoft.VSTS.Common.ClosedBy",
  "Microsoft.VSTS.Build.IntegrationBuild", "Microsoft.VSTS.Build.FoundIn",
  "Philips.Common.FoundInRelease",
]);

// PAT is injected at build time via GitHub Actions (SST_ADO_PAT secret).
// The placeholder below is replaced during the CI build step.
const INJECTED_PAT = "%%SST_ADO_PAT%%";

// ── Helpers ───────────────────────────────────────────────────────────────────

function log(line) {
  els.log.textContent += line + "\n";
  els.log.scrollTop = els.log.scrollHeight;
}

function setStatus(text, kind, spinning) {
  els.status.className = "status" + (kind ? " " + kind : "");
  els.statusText.textContent = text;
  els.spin.style.display = spinning ? "" : "none";
}

function isPhilipsEmail(v) {
  return /^[^\s@]+@philips\.com$/i.test(v.trim());
}

function refreshRunState() {
  if (!pyReady) return;
  const yearOk  = /^\d{4}$/.test(els.year.value.trim());
  const doSticr = els.createSticr.checked;
  const emailOk = !doSticr || isPhilipsEmail(els.userEmail.value);
  const ok = els.prevFile.files.length && updBytes && selectedMonth && yearOk
          && emailOk && !verBlocking;
  els.run.disabled = !ok;
  if (ok) els.run.textContent = "Generate document";

  els.emailReveal.style.display = doSticr ? "" : "none";
}

// ── Month grid ────────────────────────────────────────────────────────────────

els.monthGrid.addEventListener("click", e => {
  const btn = e.target.closest(".month-btn");
  if (!btn) return;
  els.monthGrid.querySelectorAll(".month-btn").forEach(b => b.classList.remove("selected"));
  btn.classList.add("selected");
  selectedMonth = btn.dataset.month;
  refreshRunState();
});

els.year.addEventListener("input", refreshRunState);

// ── Email validation ──────────────────────────────────────────────────────────
els.userEmail.addEventListener("input", () => {
  const v = els.userEmail.value;
  if (v && !isPhilipsEmail(v)) {
    els.emailErr.style.display = "block";
  } else {
    els.emailErr.style.display = "none";
  }
  refreshRunState();
});

// ── STICR / qual registry toggles ────────────────────────────────────────────
els.createSticr.addEventListener("change", refreshRunState);
els.createQualReg.addEventListener("change", refreshRunState);

// ── Upload zones ──────────────────────────────────────────────────────────────

function wireZone(zoneId, input, chosenEl, onBytes) {
  const zone = document.getElementById(zoneId);

  input.addEventListener("change", () => {
    const f = input.files[0];
    if (!f) return;
    f.arrayBuffer().then(buf => {
      onBytes(new Uint8Array(buf), f.name);
      chosenEl.textContent = f.name;
      chosenEl.style.display = "block";
      refreshRunState();
    });
  });

  zone.addEventListener("dragover",  e => { e.preventDefault(); zone.classList.add("drag-over"); });
  zone.addEventListener("dragleave", ()  => zone.classList.remove("drag-over"));
  zone.addEventListener("drop", e => {
    e.preventDefault();
    zone.classList.remove("drag-over");
    const f = e.dataTransfer.files[0];
    if (!f) return;
    f.arrayBuffer().then(buf => {
      onBytes(new Uint8Array(buf), f.name);
      chosenEl.textContent = f.name;
      chosenEl.style.display = "block";
      refreshRunState();
    });
  });
}

wireZone("zone-prev", els.prevFile, els.chosenPrev, () => refreshRunState());
wireZone("zone-upd",  els.updFile,  els.chosenUpd,  (bytes, name) => {
  updBytes    = bytes;
  updFileName = name;
  runVersionChecks(bytes);
});

els.verDismiss.addEventListener("click", () => {
  els.verBanner.style.display = "none";
  verBlocking = false;
  refreshRunState();
});

// ── Version checks ────────────────────────────────────────────────────────────

function runVersionChecks(bytes) {
  const text  = new TextDecoder().decode(bytes);
  const lines = text.split(/\r?\n/);

  const edgeLines = lines.filter(l => /edge/i.test(l) && /version|build/i.test(l));
  const msrtLines = lines.filter(l => /windows malicious software removal tool/i.test(l) && /v\d+\.\d+/i.test(l));

  if (!edgeLines.length && !msrtLines.length) {
    els.verBanner.style.display = "none";
    verBlocking = false;
    refreshRunState();
    return;
  }

  const issuesMap = new Map();  // summary → example line (deduplicates identical messages)

  function addIssue(summary, line) {
    if (!issuesMap.has(summary)) issuesMap.set(summary, line.trim());
  }

  // ── Edge: Version major must match Build major; all builds must be consistent ──
  if (edgeLines.length) {
    const buildVersions = new Set(
      edgeLines.map(l => l.match(/build\s+([\d.]+)/i)?.[1]).filter(Boolean)
    );

    for (const line of edgeLines) {
      const verMatch    = line.match(/version\s+(\d+)/i);
      const buildMatch  = line.match(/build\s+([\d.]+)/i);
      const statedVer   = verMatch?.[1];
      const statedBuild = buildMatch?.[1];

      if (statedVer && statedBuild) {
        const buildMajor = statedBuild.split(".")[0];
        if (statedVer !== buildMajor) {
          addIssue(
            `Edge: "Version ${statedVer}" doesn't match "Build ${statedBuild}" — should be Version ${buildMajor}`,
            line
          );
        }
      }
    }

    if (buildVersions.size > 1) {
      addIssue(
        `Edge: build numbers are inconsistent across lines — ${[...buildVersions].join(", ")}`,
        edgeLines[0]
      );
    }
  }

  // ── MSRT: all version numbers must be consistent across lines ──
  if (msrtLines.length) {
    const msrtVersions = new Set(
      msrtLines.map(l => l.match(/v(5\.\d+)/i)?.[1]).filter(Boolean)
    );

    if (msrtVersions.size > 1) {
      addIssue(
        `MSRT: inconsistent versions across lines — ${[...msrtVersions].join(", ")}. All occurrences should match.`,
        msrtLines[0]
      );
    }
  }

  // ── Render ──
  if (issuesMap.size === 0) {
    els.verBanner.style.display = "none";
    verBlocking = false;
    refreshRunState();
    return;
  }

  verBlocking = true;
  els.verBanner.style.display = "block";
  els.verBanner.classList.add("ver-error");
  refreshRunState();

  const items = [...issuesMap.entries()].map(([summary, detail]) =>
    `<li><strong>${escapeHtml(summary)}</strong><br><code style="font-size:0.78rem;font-weight:500">${escapeHtml(detail)}</code></li>`
  ).join("");
  els.verBody.innerHTML = `<ul>${items}</ul><p style="margin-top:8px;font-size:0.8rem">Fix the file and re-upload, or dismiss to proceed anyway.</p>`;
}

function escapeHtml(s) {
  return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}

// ── Pyodide boot ──────────────────────────────────────────────────────────────

async function boot() {
  // Build templateFields from the script-tag global (no fetch needed)
  if (window.STICR_TEMPLATE_FIELDS) {
    templateFields = Object.entries(window.STICR_TEMPLATE_FIELDS)
      .filter(([k, v]) => !STICR_EXCLUDED.has(k) && v !== null)
      .map(([k, v]) => ({ op: "add", path: `/fields/${k}`, value: v }));
  }

  const sticrTemplateHtml = window.STICR_TEMPLATE_HTML || null;

  try {
    setStatus("Loading Python environment (first load downloads ~20 MB)...", "", true);
    pyodide = await loadPyodide({
      indexURL: PYODIDE_INDEX,
      stdout: log,
      stderr: log,
    });

    setStatus("Loading packages (pandas, lxml)...", "", true);
    await pyodide.loadPackage(["micropip", "pandas", "lxml"]);

    setStatus("Installing python-docx and python-dateutil...", "", true);
    const micropip = pyodide.pyimport("micropip");
    await micropip.install(["python-docx", "python-dateutil"]);

    try { pyodide.FS.mkdir(WORK); } catch { /* already exists */ }

    // Write the STICR HTML template into Pyodide's filesystem so Python can open it
    if (sticrTemplateHtml) {
      pyodide.FS.writeFile(WORK + "/sticr_template.html", sticrTemplateHtml, { encoding: "utf8" });
    }

    pyReady = true;
    setStatus("Ready. Choose your files and click Generate.", "ok", false);
    refreshRunState();
  } catch (e) {
    console.error(e);
    setStatus("Failed to load the Python environment: " + (e.message || e), "err", false);
    log(String(e.stack || e));
  }
}

// ── Generate ──────────────────────────────────────────────────────────────────

async function generate() {
  els.run.disabled = true;
  els.download.style.display = "none";
  els.downloadCsv.style.display = "none";
  els.log.textContent = "";

  const month       = selectedMonth;
  const year        = els.year.value.trim();
  const updExt      = updFileName.toLowerCase().endsWith(".docx") ? "docx" : "txt";
  const userEmail   = els.userEmail.value.trim();
  const pat         = INJECTED_PAT;
  const doSticr     = els.createSticr.checked;
  const doQualReg   = els.createQualReg.checked;

  try {
    setStatus("Reading uploaded files...", "", true);
    const prevBytes = new Uint8Array(await els.prevFile.files[0].arrayBuffer());

    const prevPath = WORK + "/previous.docx";
    const updPath  = WORK + "/update." + updExt;
    pyodide.FS.writeFile(prevPath, prevBytes);
    pyodide.FS.writeFile(updPath,  updBytes);

    const env = {
      SST_MONTH:         month,
      SST_YEAR:          year,
      PREVIOUS_SST_PATH: prevPath,
      UPDATE_FILE_PATH:  updPath,
      OUTPUT_DIRECTORY:  OUT_DIR,
      SST_USER_EMAIL:    userEmail,
    };
    pyodide.runPython(`
import os, json
for k, v in json.loads(${JSON.stringify(JSON.stringify(env))}).items():
    os.environ[k] = v
`);

    setStatus("Generating document...", "", true);
    await pyodide.runPythonAsync(window.SST_PYTHON);

    const outName  = `PIC iX_Security_Status_Table_${month}_${year}.docx`;
    const outBytes = pyodide.FS.readFile(OUT_DIR + "/" + outName);

    const blob = new Blob([outBytes], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const url = URL.createObjectURL(blob);
    els.download.href        = url;
    els.download.download    = outName;
    els.download.textContent = "Download " + outName;
    els.download.style.display = "block";

    if (doQualReg) {
      try {
        const csvText = pyodide.globals.get("qual_registry_csv_output");
        if (csvText) {
          const csvBlob = new Blob([csvText], { type: "text/csv" });
          const csvUrl  = URL.createObjectURL(csvBlob);
          const csvName = `Qualification_Registry_${month}_${year}.csv`;
          els.downloadCsv.href        = csvUrl;
          els.downloadCsv.download    = csvName;
          els.downloadCsv.textContent = "Download " + csvName;
          els.downloadCsv.style.display = "block";
        }
      } catch (e) {
        log("CSV generation error: " + e);
      }
    }

    if (doSticr) {
      await createSticrs(pat, userEmail);
    } else {
      setStatus("Done. Your document is ready.", "ok", false);
    }
  } catch (e) {
    console.error(e);
    const msg    = String(e.message || e);
    const valErr = msg.match(/(?:ValueError|KeyError|FileNotFoundError|Exception):.*$/m);
    setStatus(
      "Generation failed: " + (valErr ? valErr[0] : msg.split("\n").slice(-3).join("\n")),
      "err", false
    );
    log(msg);
  } finally {
    refreshRunState();
  }
}

// ── STICR creation ────────────────────────────────────────────────────────────

async function createSticrs(pat, userEmail) {
  const organization = "PhilipsMA";
  const project      = "Sandbox";

  let sticrData;
  try {
    const raw = pyodide.globals.get("sticr_json_output");
    sticrData = JSON.parse(raw);
  } catch (e) {
    setStatus("STICR data not found — document saved but STICRs not created.", "warn", false);
    log("STICR data error: " + e);
    return;
  }

  if (!sticrData || !sticrData.length) {
    setStatus("Done. No STICRs to create.", "ok", false);
    return;
  }

  const createUrl = `https://dev.azure.com/${organization}/${project}/_apis/wit/workitems/$STICR?api-version=7.1`;
  const headers   = {
    "Content-Type": "application/json-patch+json",
    "Authorization": "Basic " + btoa(":" + pat),
  };

  let created = 0;
  let failed  = 0;
  const createdItems = [];  // { id, title, url }

  setStatus(`Creating ${sticrData.length} STICR(s) in Azure DevOps...`, "", true);

  for (const item of sticrData) {
    const body = [
      { op: "add", path: "/fields/System.Title",         value: item.title },
      { op: "add", path: "/fields/System.Description",   value: item.html },
      { op: "add", path: "/fields/System.AreaPath",       value: project },
      { op: "add", path: "/fields/System.TeamProject",    value: project },
      { op: "add", path: "/fields/System.IterationPath",  value: project + "\\Common" },
      { op: "add", path: "/fields/System.AssignedTo",     value: userEmail },
      ...templateFields,
    ];

    try {
      const resp = await fetch(createUrl, { method: "POST", headers, body: JSON.stringify(body) });
      if (!resp.ok) {
        const errText = await resp.text();
        log(`FAILED (${resp.status}): ${item.title}\n${errText}`);
        failed++;
      } else {
        const wi = await resp.json();
        const wiUrl = `https://dev.azure.com/${organization}/${project}/_workitems/edit/${wi.id}`;
        log(`Created STICR ${wi.id}: ${item.title}`);
        createdItems.push({ id: wi.id, title: item.title, url: wiUrl });
        created++;
      }
    } catch (e) {
      log(`NETWORK ERROR for "${item.title}": ${e}`);
      failed++;
    }
  }

  if (failed === 0) {
    setStatus(`Done. Document ready + ${created} STICR(s) created.`, "ok", false);
  } else {
    setStatus(`Document ready. ${created} STICR(s) created, ${failed} failed — see log.`, "warn", false);
  }

  if (createdItems.length > 0) {
    showSticrPopup(createdItems);
  }
}

function showSticrPopup(items) {
  const existing = document.getElementById("sticr-popup-overlay");
  if (existing) existing.remove();

  const rows = items.map(item =>
    `<li style="margin-bottom:10px;">
      <a href="${escapeHtml(item.url)}" target="_blank" rel="noopener"
         style="font-weight:700;color:var(--pink-dark);text-decoration:none;font-size:0.92rem;">
        #${item.id}
      </a>
      <span style="color:var(--text-soft);font-size:0.82rem;margin-left:6px;">${escapeHtml(item.title)}</span>
    </li>`
  ).join("");

  const overlay = document.createElement("div");
  overlay.id = "sticr-popup-overlay";
  overlay.style.cssText = `
    position:fixed;inset:0;background:rgba(74,25,66,0.45);z-index:1000;
    display:flex;align-items:center;justify-content:center;padding:20px;
  `;

  overlay.innerHTML = `
    <div style="
      background:#fff;border-radius:20px;padding:28px 32px;max-width:560px;width:100%;
      box-shadow:0 12px 48px rgba(219,39,119,0.22);border:1.5px solid var(--pink-mid);
      max-height:80vh;overflow-y:auto;
    ">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;">
        <strong style="font-size:1.05rem;color:var(--pink-dark);">
          STICRs created (${items.length})
        </strong>
        <button id="sticr-popup-close" style="
          background:none;border:none;cursor:pointer;font-size:1.1rem;
          color:var(--text-soft);padding:2px 6px;font-weight:800;line-height:1;
        " title="Close">&#x2715;</button>
      </div>
      <ul style="list-style:none;padding:0;margin:0 0 16px 0;">${rows}</ul>
      <p style="font-size:0.77rem;color:var(--text-soft);margin:0;">
        Click any ID to open the STICR in Azure DevOps.
      </p>
    </div>
  `;

  document.body.appendChild(overlay);

  document.getElementById("sticr-popup-close").addEventListener("click", () => overlay.remove());
  overlay.addEventListener("click", e => { if (e.target === overlay) overlay.remove(); });
}

els.run.addEventListener("click", generate);

boot();

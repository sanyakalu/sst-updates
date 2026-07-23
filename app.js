/* SST Update Generator — glue between the UI and Pyodide. */

const PYODIDE_INDEX = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";
const WORK    = "/work";
const OUT_DIR = WORK + "/output";

// ── DOM refs ──────────────────────────────────────────────────────────────────

const els = {
  prevFile:   document.getElementById("prevFile"),
  updFile:    document.getElementById("updFile"),
  chosenPrev: document.getElementById("chosen-prev"),
  chosenUpd:  document.getElementById("chosen-upd"),
  monthGrid:  document.getElementById("month-grid"),
  year:       document.getElementById("year"),
  run:        document.getElementById("run"),
  status:     document.getElementById("status"),
  statusText: document.getElementById("statusText"),
  spin:       document.getElementById("spin"),
  download:   document.getElementById("download"),
  log:        document.getElementById("log"),
  verBanner:  document.getElementById("ver-banner"),
  verBody:    document.getElementById("ver-body"),
  verDismiss: document.getElementById("ver-dismiss"),
};

// ── State ─────────────────────────────────────────────────────────────────────

let pyodide       = null;
let pyReady       = false;
let selectedMonth = null;
let updBytes      = null;
let updFileName   = "update.txt";
let verBlocking   = false;  // true when version check found issues

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

function refreshRunState() {
  if (!pyReady) return;
  const yearOk = /^\d{4}$/.test(els.year.value.trim());
  const ok = els.prevFile.files.length && updBytes && selectedMonth && yearOk && !verBlocking;
  els.run.disabled = !ok;
  if (ok) els.run.textContent = "Generate document";
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
  els.log.textContent = "";

  const month  = selectedMonth;
  const year   = els.year.value.trim();
  const updExt = updFileName.toLowerCase().endsWith(".docx") ? "docx" : "txt";

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

    setStatus("Done. Your document is ready.", "ok", false);
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

els.run.addEventListener("click", generate);

boot();

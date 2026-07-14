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
};

// ── State ─────────────────────────────────────────────────────────────────────

let pyodide       = null;
let pyReady       = false;
let selectedMonth = null;
let updBytes      = null;
let updFileName   = "update.txt";

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
  const ok = els.prevFile.files.length && updBytes && selectedMonth && yearOk;
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
});

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

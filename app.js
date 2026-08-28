/* SST Update Generator — glue between the UI and Pyodide. */

const PYODIDE_INDEX = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";
const WORK    = "/work";
const OUT_DIR = WORK + "/output";
const INJECTED_PAT = "%%SST_ADO_PAT%%";
const DVR_TEMPLATE_URL = "https://raw.githubusercontent.com/sanyakalu/sst-updates/main/DVR_TEMPLATE.docx";
const MONTH_NUM = {Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};

// ── DOM refs ──────────────────────────────────────────────────────────────────

const els = {
  prevFile:         document.getElementById("prevFile"),
  updFile:          document.getElementById("updFile"),
  chosenPrev:       document.getElementById("chosen-prev"),
  chosenUpd:        document.getElementById("chosen-upd"),
  monthGrid:        document.getElementById("month-grid"),
  year:             document.getElementById("year"),
  userEmail:        document.getElementById("userEmail"),
  emailErr:         document.getElementById("emailErr"),
  emailReveal:      document.getElementById("emailReveal"),
  doMonthly:        document.getElementById("doMonthly"),
  doSticr:          document.getElementById("doSticr"),
  doQualReg:        document.getElementById("doQualReg"),
  updSection:       document.getElementById("upd-section"),
  monthYearSection: document.getElementById("month-year-section"),
  sticrWarning:     document.getElementById("sticr-warning"),
  sticrProjectRow:  document.getElementById("sticr-project-row"),
  btnProjTest:      document.getElementById("btn-proj-test"),
  btnProjPic:       document.getElementById("btn-proj-pic"),
  run:              document.getElementById("run"),
  status:           document.getElementById("status"),
  statusText:       document.getElementById("statusText"),
  spin:             document.getElementById("spin"),
  download:         document.getElementById("download"),
  downloadCsv:      document.getElementById("download-csv"),
  log:              document.getElementById("log"),
  verBanner:        document.getElementById("ver-banner"),
  verBody:          document.getElementById("ver-body"),
  verDismiss:       document.getElementById("ver-dismiss"),
  doDvr:            document.getElementById("doDvr"),
  downloadDvr:      document.getElementById("download-dvr"),
};

// ── State ─────────────────────────────────────────────────────────────────────

let pyodide       = null;
let pyReady       = false;
let selectedMonth = null;
let updBytes      = null;
let updFileName   = "update.txt";
let verBlocking   = false;  // true when version check found issues
let sticrProject  = "Sandbox";  // toggled by Test/PiC buttons

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

// ── Refresh run state ─────────────────────────────────────────────────────────

function refreshRunState() {
  const doDvr    = els.doDvr.checked;
  const doMonthly = els.doMonthly.checked;
  const doSticr   = els.doSticr.checked;
  const doQualReg = els.doQualReg.checked;

  const fieldPrev    = document.getElementById("field-prev");
  const fieldMonthly = document.getElementById("field-monthly");
  const fieldSticr   = document.getElementById("field-sticr");
  const fieldQualreg = document.getElementById("field-qualreg");

  // DVR mode: hide all SST inputs, show only month/year picker
  if (doDvr) {
    fieldPrev.style.display    = "none";
    fieldMonthly.style.display = "none";
    fieldSticr.style.display   = "none";
    fieldQualreg.style.display = "none";
    els.updSection.style.display       = "none";
    els.monthYearSection.style.display = "block";
    els.sticrWarning.style.display     = "none";
    els.sticrProjectRow.style.display  = "none";
    els.emailReveal.style.display      = "none";
    els.emailErr.style.display         = "none";
    els.verBanner.style.display        = "none";
    els.run.textContent = "Generate DVR";
    if (!pyReady) { els.run.disabled = true; return; }
    const yearOk = /^\d{4}$/.test(els.year.value.trim());
    els.run.disabled = !(!!selectedMonth && yearOk);
    return;
  }

  // Restore SST field visibility when DVR is off
  fieldPrev.style.display    = "";
  fieldMonthly.style.display = "";
  fieldSticr.style.display   = "";
  fieldQualreg.style.display = "";

  // Show/hide conditional sections (no pyReady gate — layout updates immediately)
  els.updSection.style.display       = doMonthly ? "block" : "none";
  els.monthYearSection.style.display = doMonthly ? "block" : "none";

  // STICR warning: shown when STICR on but monthly off
  els.sticrWarning.style.display = (doSticr && !doMonthly) ? "block" : "none";

  // Project selector: shown whenever STICR is on
  els.sticrProjectRow.style.display = doSticr ? "flex" : "none";

  // Email reveal: shown when both STICR and monthly are on (independent of pyReady)
  els.emailReveal.style.display = (doSticr && doMonthly) ? "block" : "none";

  // Email error hint
  if (els.userEmail.value && !isPhilipsEmail(els.userEmail.value)) {
    els.emailErr.style.display = "block";
  } else {
    els.emailErr.style.display = "none";
  }

  // Determine run button label
  const noneToggled = !doMonthly && !doSticr && !doQualReg;
  let label;
  if (noneToggled) {
    label = "Select at least one option";
  } else if (doMonthly && doSticr && doQualReg) {
    label = "Generate document + STICRs + qualification CSV";
  } else if (doMonthly && doSticr) {
    label = "Generate document + create STICRs";
  } else if (doMonthly && doQualReg) {
    label = "Generate document + qualification CSV";
  } else if (doMonthly) {
    label = "Generate document";
  } else if (doQualReg) {
    label = "Generate qualification CSV";
  } else {
    // doSticr only (doMonthly is off, warning is shown)
    label = "Select at least one option";
  }

  if (!pyReady) {
    els.run.disabled = true;
    return;
  }

  // Compute ok
  let ok = els.prevFile.files.length > 0 && !verBlocking;

  if (noneToggled || (doSticr && !doMonthly && !doQualReg)) {
    ok = false;
  } else {
    if (doMonthly) {
      const yearOk = /^\d{4}$/.test(els.year.value.trim());
      ok = ok && !!updBytes && !!selectedMonth && yearOk;
    }
    if (doSticr && doMonthly) {
      ok = ok && isPhilipsEmail(els.userEmail.value);
    }
  }

  els.run.disabled = !ok;
  els.run.textContent = label;
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
els.userEmail.addEventListener("input", refreshRunState);

// ── Toggle wiring ─────────────────────────────────────────────────────────────
els.doDvr.addEventListener("change",     refreshRunState);
els.doMonthly.addEventListener("change", refreshRunState);
els.doSticr.addEventListener("change",   refreshRunState);
els.doQualReg.addEventListener("change", refreshRunState);

// ── STICR project selector ────────────────────────────────────────────────────
els.btnProjTest.addEventListener("click", () => {
  sticrProject = "Sandbox";
  els.btnProjTest.classList.add("active");
  els.btnProjPic.classList.remove("active");
});
els.btnProjPic.addEventListener("click", () => {
  sticrProject = "Philips.PIC";
  els.btnProjPic.classList.add("active");
  els.btnProjTest.classList.remove("active");
});

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

// ── DVR ADO data fetch ────────────────────────────────────────────────────────

async function fetchDvrAdo(pat, month, year) {
  const organization = "PhilipsMA";
  const project      = "Philips.PIC";
  const authHeader = {
    "Authorization": "Basic " + btoa(":" + pat),
    "Content-Type": "application/json",
  };
  const getHeaders = { "Authorization": "Basic " + btoa(":" + pat) };
  const monthNum   = MONTH_NUM[month];
  const prevMonth  = monthNum > 1 ? monthNum - 1 : 12;
  const prevYear   = monthNum > 1 ? parseInt(year) : parseInt(year) - 1;
  const monthStart = `${prevYear}-${String(prevMonth).padStart(2,'0')}-16`;
  const monthEnd   = `${year}-${String(monthNum).padStart(2,'0')}-15`;
  const areaPath   = "Philips.PIC\\SysEng - ESS";
  const wiqlUrl    = `https://dev.azure.com/${organization}/${project}/_apis/wit/wiql?api-version=7.1`;

  const keywords = [
    "Crowdstrike","VMWare","OS Security","SQL Server",
    "Symantec","TrendMicro","Trellix","VMware","Nutanix","Hyper-V",
    "Microsoft Security Updates",
  ];
  const kwClauses = keywords.map(k => `[System.Title] CONTAINS '${k}'`).join(" OR ");

  // Test plans
  const planResp = await fetch(wiqlUrl, {
    method: "POST", headers: authHeader,
    body: JSON.stringify({ query:
      `SELECT [System.Id],[System.Title],[System.CreatedDate] FROM WorkItems ` +
      `WHERE [System.WorkItemType]='Test Plan' ` +
      `AND [System.AreaPath] UNDER '${areaPath}' ` +
      `AND [System.CreatedDate]>='${monthStart}' ` +
      `AND [System.CreatedDate]<='${monthEnd}' ` +
      `AND (${kwClauses}) ORDER BY [System.CreatedDate] DESC`
    }),
  });
  const planIds = ((await planResp.json()).workItems || []).map(w => w.id);
  log(`DVR: found ${planIds.length} test plan(s)`);

  const testRows = [];
  for (const planId of planIds) {
    const plan = await (await fetch(
      `https://dev.azure.com/${organization}/${project}/_apis/testplan/plans/${planId}?api-version=7.1`,
      { headers: getHeaders }
    )).json();
    const runs = ((await (await fetch(
      `https://dev.azure.com/${organization}/${project}/_apis/test/runs?planId=${planId}&api-version=7.1`,
      { headers: getHeaders }
    )).json()).value) || [];
    for (const run of runs) {
      const results = ((await (await fetch(
        `https://dev.azure.com/${organization}/${project}/_apis/test/runs/${run.id}/results?api-version=7.1`,
        { headers: getHeaders }
      )).json()).value) || [];
      for (const r of results) {
        testRows.push([
          planId, plan.name || "", run.name || "", run.id,
          run.state || "N/A", r.outcome || "N/A",
          r.testCaseTitle || "N/A", r.testCase?.id || "N/A",
        ]);
      }
    }
  }
  log(`DVR: fetched ${testRows.length} test result(s)`);

  // STICRs
  const MONTH_NAMES_FULL = {1:"January",2:"February",3:"March",4:"April",5:"May",6:"June",7:"July",8:"August",9:"September",10:"October",11:"November",12:"December"};
  const sticrResp = await fetch(wiqlUrl, {
    method: "POST", headers: authHeader,
    body: JSON.stringify({ query:
      `SELECT [System.Id],[System.Title] FROM WorkItems ` +
      `WHERE [System.WorkItemType]='STICR' ` +
      `AND [System.AreaPath] UNDER '${areaPath}' ` +
      `AND [System.TeamProject]='Philips.PIC' ` +
      `AND [System.Title] CONTAINS 'Microsoft Security Update' ` +
      `AND [System.Title] CONTAINS '${MONTH_NAMES_FULL[monthNum]}' ` +
      `AND [System.Title] CONTAINS '${year}' ` +
      `ORDER BY [System.CreatedDate] DESC`
    }),
  });
  const sticrIds = ((await sticrResp.json()).workItems || []).map(w => w.id);
  log(`DVR: found ${sticrIds.length} STICR(s)`);

  const sticrItems = [];
  for (const sid of sticrIds) {
    const item = await (await fetch(
      `https://dev.azure.com/${organization}/${project}/_apis/wit/workitems/${sid}?$expand=fields&api-version=7.1`,
      { headers: getHeaders }
    )).json();
    sticrItems.push({
      id:          sid,
      description: item.fields?.["System.Description"] || "",
    });
  }

  return { testRows, sticrItems };
}

// ── Generate ──────────────────────────────────────────────────────────────────

async function generate() {
  els.run.disabled = true;
  els.download.style.display    = "none";
  els.downloadCsv.style.display = "none";
  els.downloadDvr.style.display = "none";
  els.log.textContent = "";

  const doMonthly = els.doMonthly.checked;
  const doSticr   = els.doSticr.checked;
  const doQualReg = els.doQualReg.checked;

  const month     = selectedMonth;
  const year      = els.year.value.trim();
  const updExt    = updFileName.toLowerCase().endsWith(".docx") ? "docx" : "txt";
  const userEmail = els.userEmail.value.trim();
  const pat       = INJECTED_PAT;

  try {
    // ── DVR branch ──────────────────────────────────────────────────────────
    if (els.doDvr.checked) {
      const month    = selectedMonth;
      const year     = els.year.value.trim();
      const monthNum = MONTH_NUM[month];
      const pat      = INJECTED_PAT;

      setStatus("Fetching DVR template from repo...", "", true);
      const tplResp = await fetch(DVR_TEMPLATE_URL);
      if (!tplResp.ok) throw new Error(`Failed to fetch DVR template: ${tplResp.status}`);
      const tplBytes = new Uint8Array(await tplResp.arrayBuffer());
      try { pyodide.FS.mkdir(OUT_DIR); } catch { /* already exists */ }
      pyodide.FS.writeFile("/work/dvr_template.docx", tplBytes);

      setStatus("Fetching Azure DevOps data...", "", true);
      const { testRows, sticrItems } = await fetchDvrAdo(pat, month, year);

      pyodide.globals.set("dvr_test_rows_json",    JSON.stringify(testRows));
      pyodide.globals.set("dvr_sticr_items_json",  JSON.stringify(sticrItems));

      const dvrEnv = { DVR_MONTH_NUM: String(monthNum), DVR_YEAR: year, OUTPUT_DIRECTORY: OUT_DIR };
      pyodide.runPython(`
import os, json
for k, v in json.loads(${JSON.stringify(JSON.stringify(dvrEnv))}).items():
    os.environ[k] = v
`);

      setStatus("Generating DVR document...", "", true);
      await pyodide.runPythonAsync(window.DVR_PYTHON);

      const outName  = pyodide.globals.get("dvr_output_filename");
      const outBytes = pyodide.FS.readFile(OUT_DIR + "/" + outName);
      const blob = new Blob([outBytes], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      const url = URL.createObjectURL(blob);
      els.downloadDvr.href        = url;
      els.downloadDvr.download    = outName;
      els.downloadDvr.textContent = "Download " + outName;
      els.downloadDvr.style.display = "block";
      setStatus("Done. Your DVR is ready.", "ok", false);
      return;
    }
    // ── End DVR branch ──────────────────────────────────────────────────────

    setStatus("Reading uploaded files...", "", true);
    const prevBytes = new Uint8Array(await els.prevFile.files[0].arrayBuffer());

    const prevPath = WORK + "/previous.docx";
    pyodide.FS.writeFile(prevPath, prevBytes);

    // Always set PREVIOUS_SST_PATH (used by both monthly and qual reg)
    const baseEnv = {
      PREVIOUS_SST_PATH: prevPath,
    };
    pyodide.runPython(`
import os, json
for k, v in json.loads(${JSON.stringify(JSON.stringify(baseEnv))}).items():
    os.environ[k] = v
`);

    // ── Monthly SST update ────────────────────────────────────────────────────
    if (doMonthly) {
      const updPath = WORK + "/update." + updExt;
      pyodide.FS.writeFile(updPath, updBytes);

      const monthlyEnv = {
        SST_MONTH:        month,
        SST_YEAR:         year,
        UPDATE_FILE_PATH: updPath,
        OUTPUT_DIRECTORY: OUT_DIR,
        SST_USER_EMAIL:   userEmail,
      };
      pyodide.runPython(`
import os, json
for k, v in json.loads(${JSON.stringify(JSON.stringify(monthlyEnv))}).items():
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
    }

    // ── STICRs ────────────────────────────────────────────────────────────────
    if (doSticr && doMonthly) {
      await createSticrs(pat, userEmail);
    }

    // ── Qualification CSV ─────────────────────────────────────────────────────
    if (doQualReg) {
      // SST_MONTH and SST_YEAR may be needed by QUAL_PYTHON even if monthly is off;
      // set them from the year field (best effort — required if doMonthly is also on)
      const qualEnv = {
        SST_MONTH: selectedMonth || "",
        SST_YEAR:  els.year.value.trim(),
      };
      pyodide.runPython(`
import os, json
for k, v in json.loads(${JSON.stringify(JSON.stringify(qualEnv))}).items():
    os.environ[k] = v
`);

      setStatus("Generating qualification CSV...", "", true);
      await pyodide.runPythonAsync(window.QUAL_PYTHON);

      const csvContent = pyodide.globals.get("qual_registry_csv_output");
      if (csvContent) {
        const csvBlob = new Blob([csvContent], { type: "text/csv" });
        const csvUrl  = URL.createObjectURL(csvBlob);
        const csvName = selectedMonth
          ? `Qualification_Registry_${selectedMonth}_${els.year.value.trim()}.csv`
          : `Qualification_Registry_${els.year.value.trim()}.csv`;
        els.downloadCsv.href        = csvUrl;
        els.downloadCsv.download    = csvName;
        els.downloadCsv.textContent = "Download " + csvName;
        els.downloadCsv.style.display = "block";
      } else {
        log("Warning: qual_registry_csv_output was empty or not set.");
      }
    }

    // ── Final status ──────────────────────────────────────────────────────────
    if (!doSticr || !doMonthly) {
      // createSticrs sets its own final status; only set here if STICRs weren't run
      setStatus("Done. Your output is ready.", "ok", false);
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
  const project      = sticrProject;

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

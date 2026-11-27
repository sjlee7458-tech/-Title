importScripts("https://cdn.jsdelivr.net/pyodide/v0.26.2/full/pyodide.js");

let pyodidePromise = null;

async function initPyodide() {
  if (!pyodidePromise) {
    pyodidePromise = (async () => {
      const pyodide = await loadPyodide();
      await pyodide.loadPackage(["numpy", "pandas", "micropip"]);
      const micropip = pyodide.pyimport("micropip");
      await micropip.install("openpyxl");
      const resp = await fetch("./model.py");
      const code = await resp.text();
      await pyodide.runPythonAsync(code);
      return pyodide;
    })();
  }
  return pyodidePromise;
}

self.onmessage = async (event) => {
  const { id, type, payload } = event.data;
  try {
    const pyodide = await initPyodide();
    if (type === "PING") {
      self.postMessage({ id, ok: true, data: { ready: true } });
      return;
    }

    if (type === "LOAD_EXCEL") {
      const bytes = new Uint8Array(payload.buffer);
      pyodide.globals.set("excel_bytes", pyodide.toPy(bytes));
      const jsonStr = await pyodide.runPythonAsync("load_workbook_summary(bytes(excel_bytes))");
      self.postMessage({ id, ok: true, data: JSON.parse(jsonStr) });
      return;
    }

    if (type === "RUN_ANALYSIS") {
      pyodide.globals.set("sheet_name", payload.sheetName);
      pyodide.globals.set("target_sector", payload.targetSector);
      pyodide.globals.set("demand_shock", Number(payload.demandShock || 0));
      pyodide.globals.set("shock_label", payload.shockLabel || "");
      const jsonStr = await pyodide.runPythonAsync(
        "run_io_analysis(sheet_name, target_sector, float(demand_shock), shock_label)"
      );
      self.postMessage({ id, ok: true, data: JSON.parse(jsonStr) });
    }
  } catch (err) {
    self.postMessage({ id, ok: false, error: err.message || String(err) });
  }
};

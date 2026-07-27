const form = document.querySelector("#reconcileForm");
const runButton = document.querySelector("#runButton");
const statusBox = document.querySelector("#status");
const results = document.querySelector("#results");
const downloadLink = document.querySelector("#downloadLink");
const reconciledCount = document.querySelector("#reconciledCount");
const exceptionCount = document.querySelector("#exceptionCount");

function wireFileName(inputId, labelId) {
  const input = document.querySelector(inputId);
  const label = document.querySelector(labelId);
  input.addEventListener("change", () => {
    label.textContent = input.files.length ? input.files[0].name : label.dataset.placeholder;
  });
}

function setStatus(message, state = "idle") {
  statusBox.textContent = message;
  statusBox.className = `status ${state}`;
}

function renderTable(selector, rows) {
  const table = document.querySelector(selector);
  table.innerHTML = "";
  if (!rows || rows.length === 0) {
    table.innerHTML = "<tbody><tr><td>No records</td></tr></tbody>";
    return;
  }

  const columns = Object.keys(rows[0]);
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  columns.forEach((column) => {
    const th = document.createElement("th");
    th.textContent = column.replace(/([A-Z])/g, " $1");
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    columns.forEach((column) => {
      const td = document.createElement("td");
      td.textContent = row[column];
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
}

wireFileName("#payroll", "#payrollName");
wireFileName("#billing", "#billingName");

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const data = new FormData(form);
  runButton.disabled = true;
  results.classList.add("hidden");
  setStatus("Running reconciliation...", "idle");

  try {
    const response = await fetch("/reconcile", { method: "POST", body: data });
    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.error || "Reconciliation failed.");
    }

    reconciledCount.textContent = payload.reconciledCount;
    exceptionCount.textContent = payload.exceptionCount;
    downloadLink.href = payload.downloadUrl;
    renderTable("#summaryTable", payload.summary);
    renderTable("#exceptionTable", payload.exceptionSummary);
    results.classList.remove("hidden");
    setStatus(`Completed. Report generated: ${payload.file}`, "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    runButton.disabled = false;
  }
});


(function () {
  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("file-input");
  const browseBtn = document.getElementById("browse-btn");
  const fileLabel = document.getElementById("file-label");
  const previewCard = document.getElementById("preview-card");
  const previewBody = document.getElementById("preview-body");
  const columnName = document.getElementById("column-name");
  const statTotal = document.getElementById("stat-total");
  const statOk = document.getElementById("stat-ok");
  const statBad = document.getElementById("stat-bad");
  const downloadBtn = document.getElementById("download-btn");
  const resetBtn = document.getElementById("reset-btn");
  const statusBox = document.getElementById("status-box");

  let currentRows = [];
  let sourceName = "phones";

  function setStatus(message, kind) {
    if (!statusBox) return;
    statusBox.hidden = !message;
    statusBox.className = "status " + (kind || "");
    statusBox.textContent = message || "";
  }

  function sheetToRows(workbook) {
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      defval: "",
      blankrows: true,
    });
  }

  function safeFileName(name) {
    return String(name || "phones")
      .replace(/[\\/:*?"<>|]+/g, "_")
      .replace(/\s+/g, "_")
      .slice(0, 80) || "phones";
  }

  function renderPreview(result, fileName) {
    currentRows = result.rows;
    sourceName = safeFileName(fileName.replace(/\.[^.]+$/, "") || "phones");
    columnName.textContent = result.phoneColumn.header || "לא זוהתה";

    const okCount = result.rows.filter((row) => row.ok).length;
    statTotal.textContent = String(result.rows.length);
    statOk.textContent = String(okCount);
    statBad.textContent = String(result.rows.length - okCount);
    downloadBtn.disabled = okCount === 0;
    setStatus("", "");

    previewBody.innerHTML = "";
    result.rows.forEach((row) => {
      const tr = document.createElement("tr");
      tr.className = row.ok ? "ok" : "bad";
      tr.innerHTML =
        "<td>" + row.sourceRow + "</td>" +
        "<td class=\"raw\">" + escapeHtml(row.raw) + "</td>" +
        "<td class=\"formatted\">" + escapeHtml(row.formatted || "—") + "</td>" +
        "<td>" + escapeHtml(row.ok ? (row.warning || "תקין") : row.reason) + "</td>";
      previewBody.appendChild(tr);
    });

    previewCard.hidden = false;
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function handleFile(file) {
    if (!file) return;
    fileLabel.textContent = file.name;
    dropzone.classList.add("has-file");
    setStatus("קורא את הקובץ...", "");

    file.arrayBuffer()
      .then((buffer) => XLSX.read(buffer, { type: "array", raw: false, cellText: true }))
      .then((workbook) => {
        const rows = sheetToRows(workbook);
        if (!rows.length) throw new Error("הקובץ ריק");
        renderPreview(PhoneFix.collectRows(rows), file.name);
      })
      .catch((error) => {
        setStatus("לא הצלחתי לקרוא את הקובץ: " + error.message, "bad");
      });
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.rel = "noopener";
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
  }

  function downloadOutput() {
    const validRows = currentRows.filter((row) => row.ok);
    if (!validRows.length) return;

    const filename = sourceName + "_User_list.xlsx";
    const payload = {
      filename: filename,
      rows: validRows.map(function (row) {
        return {
          phone: row.formatted,
          firstName: row.firstName || "",
          lastName: row.lastName || "",
        };
      }),
    };

    downloadBtn.disabled = true;
    setStatus("יוצר קובץ מתוך התבנית המקורית ומדביק רק Values...", "");

    fetch("/export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    })
      .then(function (response) {
        if (!response.ok) {
          return response.text().then(function (text) {
            throw new Error(text || "export failed");
          });
        }
        return response.blob();
      })
      .then(function (blob) {
        if (!blob || blob.size < 1000) throw new Error("הקובץ שנוצר ריק");
        downloadBlob(blob, filename);
        setStatus("הקובץ ירד. המספרים הודבקו כערכים בלבד בעמודת Phone number.", "ok");
      })
      .catch(function (error) {
        setStatus(
          "ההורדה נכשלה. הפעילו את המערכת דרך Start.bat ואז נסו שוב. " + error.message,
          "bad"
        );
      })
      .then(function () {
        downloadBtn.disabled = validRows.length === 0;
      });
  }

  function resetApp() {
    currentRows = [];
    fileInput.value = "";
    fileLabel.textContent = "לא נבחר קובץ";
    dropzone.classList.remove("has-file");
    previewCard.hidden = true;
    downloadBtn.disabled = true;
    setStatus("", "");
  }

  ["dragenter", "dragover"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.add("drag");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.remove("drag");
    });
  });

  dropzone.addEventListener("drop", (event) => {
    const file = event.dataTransfer.files && event.dataTransfer.files[0];
    handleFile(file);
  });

  browseBtn.addEventListener("click", (event) => {
    event.stopPropagation();
    fileInput.click();
  });
  dropzone.addEventListener("click", (event) => {
    if (event.target === browseBtn) return;
    fileInput.click();
  });
  fileInput.addEventListener("change", () => handleFile(fileInput.files[0]));
  downloadBtn.addEventListener("click", downloadOutput);
  resetBtn.addEventListener("click", resetApp);
})();

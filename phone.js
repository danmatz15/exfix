(function (root) {
  const PHONE_HEADERS = [
    "phone numbers",
    "phone number",
    "phonenumbers",
    "phonenumber",
    "phone",
    "mobile",
    "cell",
    "cellphone",
    "tel",
    "telephone",
    "טלפון",
    "נייד",
    "פלאפון",
    "מספר טלפון",
    "מספר נייד",
    "טלפון נייד",
    "מספר פלאפון",
  ];

  const NAME_HEADERS = {
    first: ["first name", "firstname", "first", "שם פרטי", "שם מלא", "שם", "name"],
    last: ["last name", "lastname", "last", "surname", "שם משפחה", "משפחה"],
  };

  function normalizeHeader(value) {
    return String(value || "")
      .replace(/\u200f|\u200e|\ufeff/g, "")
      .trim()
      .toLowerCase();
  }

  function headerScore(header) {
    const h = normalizeHeader(header);
    if (!h) return 0;
    if (h === "phone numbers" || h === "phone number") return 100;
    if (PHONE_HEADERS.includes(h)) return 80;
    if (h.includes("phone") || h.includes("טלפון") || h.includes("נייד") || h.includes("פלאפון")) {
      return 60;
    }
    return 0;
  }

  function looksLikePhone(value) {
    if (value == null || value === "") return false;
    const digits = String(value).replace(/[^\d]/g, "");
    return digits.length >= 8 && digits.length <= 15;
  }

  function expandScientific(raw) {
    const text = String(raw).trim();
    if (!/^[+-]?\d+(\.\d+)?[eE][+-]?\d+$/.test(text)) return text;
    const num = Number(text);
    if (!Number.isFinite(num)) return text;
    return num.toLocaleString("en-US", { useGrouping: false, maximumFractionDigits: 0 });
  }

  function splitCandidates(raw) {
    const text = expandScientific(raw);
    return String(text)
      .split(/[\n\r,;/|]+/)
      .map((part) => part.trim())
      .filter(Boolean);
  }

  function splitName(full) {
    const parts = String(full || "").trim().split(/\s+/).filter(Boolean);
    if (!parts.length) return { first: "", last: "" };
    if (parts.length === 1) return { first: parts[0], last: "" };
    return { first: parts[0], last: parts.slice(1).join(" ") };
  }

  function formatPhone(raw) {
    if (raw == null || String(raw).trim() === "") {
      return { ok: false, value: "", reason: "ריק" };
    }

    let digits = expandScientific(String(raw).trim()).replace(/[^\d]/g, "");
    if (!digits) {
      return { ok: false, value: "", reason: "אין ספרות" };
    }

    if (digits.startsWith("00")) digits = digits.slice(2);
    if (digits.startsWith("9720")) digits = "972" + digits.slice(4);

    if (digits.startsWith("972")) {
      if (digits.length === 12) return { ok: true, value: digits };
      if (digits.length >= 11 && digits.length <= 15) {
        return { ok: true, value: digits, warning: "אורך לא סטנדרטי" };
      }
      return { ok: false, value: digits, reason: "אורך שגוי אחרי קידומת 972" };
    }

    if (/^0\d{9}$/.test(digits)) {
      return { ok: true, value: "972" + digits.slice(1) };
    }

    if (/^5\d{8}$/.test(digits)) {
      return { ok: true, value: "972" + digits };
    }

    if (/^0[2-489]\d{7,8}$/.test(digits)) {
      return { ok: true, value: "972" + digits.slice(1) };
    }

    if (!digits.startsWith("0") && digits.length >= 10 && digits.length <= 15) {
      return { ok: true, value: digits };
    }

    return { ok: false, value: digits, reason: "לא זוהה פורמט" };
  }

  function findHeaderRow(rows) {
    const limit = Math.min(rows.length, 25);
    for (let r = 0; r < limit; r += 1) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c += 1) {
        if (headerScore(row[c]) > 0) return r;
      }
    }
    return -1;
  }

  function detectPhoneColumnByValues(rows, start) {
    const counts = [];
    for (let r = start; r < rows.length; r += 1) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c += 1) {
        counts[c] = (counts[c] || 0) + (looksLikePhone(row[c]) ? 1 : 0);
      }
    }
    let bestIndex = 0;
    for (let i = 1; i < counts.length; i += 1) {
      if ((counts[i] || 0) > (counts[bestIndex] || 0)) bestIndex = i;
    }
    return {
      index: (counts[bestIndex] || 0) > 0 ? bestIndex : 0,
      header: "עמודה " + (bestIndex + 1),
      score: counts[bestIndex] || 0,
    };
  }

  function detectNameColumn(headers, kind) {
    const wanted = NAME_HEADERS[kind] || [];
    for (let i = 0; i < headers.length; i += 1) {
      if (wanted.includes(normalizeHeader(headers[i]))) return i;
    }
    return -1;
  }

  function detectTextColumn(rows, start, phoneIndex) {
    const scores = [];
    for (let r = start; r < rows.length; r += 1) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c += 1) {
        if (c === phoneIndex) continue;
        const value = String(row[c] || "").trim();
        if (!value || looksLikePhone(value) || /^\d+$/.test(value)) continue;
        scores[c] = (scores[c] || 0) + 1;
      }
    }
    let best = -1;
    for (let i = 0; i < scores.length; i += 1) {
      if ((scores[i] || 0) > (best < 0 ? 0 : scores[best] || 0)) best = i;
    }
    return best;
  }

  function collectRows(rows) {
    const headerRow = findHeaderRow(rows);
    const start = headerRow >= 0 ? headerRow + 1 : 0;
    const headers = headerRow >= 0 ? rows[headerRow] || [] : [];

    let phoneCol;
    if (headerRow >= 0) {
      let best = { index: 0, header: "", score: 0 };
      headers.forEach((cell, index) => {
        const score = headerScore(cell);
        if (score > best.score) best = { index, header: String(cell || ""), score };
      });
      phoneCol = best.score > 0 ? best : detectPhoneColumnByValues(rows, start);
      if (!phoneCol.header) phoneCol.header = String(headers[phoneCol.index] || "PHONE NUMBERS");
    } else {
      phoneCol = detectPhoneColumnByValues(rows, start);
    }

    let firstCol = headerRow >= 0 ? detectNameColumn(headers, "first") : -1;
    let lastCol = headerRow >= 0 ? detectNameColumn(headers, "last") : -1;
    if (firstCol < 0) firstCol = detectTextColumn(rows, start, phoneCol.index);

    const seen = new Map();
    const results = [];

    for (let r = start; r < rows.length; r += 1) {
      const row = rows[r] || [];
      const raw = row[phoneCol.index];
      if (raw == null || String(raw).trim() === "") continue;
      if (headerScore(raw) > 0) continue;

      const parts = splitCandidates(raw);
      for (const part of parts) {
        if (!looksLikePhone(part) && headerScore(part) > 0) continue;
        const formatted = formatPhone(part);
        const key = formatted.value || String(part);
        let firstName = firstCol >= 0 ? String(row[firstCol] || "").trim() : "";
        let lastName = lastCol >= 0 ? String(row[lastCol] || "").trim() : "";
        if (firstName && lastCol < 0) {
          const split = splitName(firstName);
          firstName = split.first;
          lastName = split.last;
        }

        const item = {
          sourceRow: r + 1,
          raw: part,
          formatted: formatted.value,
          ok: formatted.ok,
          warning: formatted.warning || "",
          reason: formatted.reason || "",
          firstName,
          lastName,
          duplicate: false,
        };

        if (formatted.ok && seen.has(key)) {
          item.duplicate = true;
          item.ok = false;
          item.reason = "כפילות";
        } else if (formatted.ok) {
          seen.set(key, true);
        }

        results.push(item);
      }
    }

    return { phoneColumn: phoneCol, rows: results };
  }

  root.PhoneFix = {
    formatPhone,
    collectRows,
    headerScore,
  };
})(typeof window !== "undefined" ? window : globalThis);

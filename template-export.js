(function (root) {
  function xmlEscape(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function stringTag(value) {
    const text = String(value);
    const space = /^\s|\s$/.test(text) ? ' xml:space="preserve"' : "";
    return "<si><t" + space + ">" + xmlEscape(text) + "</t></si>";
  }

  function textCell(ref, index) {
    return '<c r="' + ref + '" s="1" t="s"><v>' + index + "</v></c>";
  }

  function emptyCell(ref) {
    return '<c r="' + ref + '"/>';
  }

  function base64ToUint8(base64) {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return bytes;
  }

  function addSharedString(additions, nextIndex, value) {
    additions.push(stringTag(value));
    const index = nextIndex.value;
    nextIndex.value += 1;
    return index;
  }

  function patchSharedStrings(sstXml, additions, total) {
    return sstXml
      .replace(/count="\d+"/, 'count="' + total + '"')
      .replace(/uniqueCount="\d+"/, 'uniqueCount="' + total + '"')
      .replace("</sst>", additions.join("") + "</sst>");
  }

  function replaceOrInsertRow(sheetXml, rowNum, aXml, bXml, cXml) {
    const rowRe = new RegExp("<row r=\"" + rowNum + "\"[\\s\\S]*?</row>");
    if (rowRe.test(sheetXml)) {
      return sheetXml.replace(rowRe, function (full) {
        let next = full.replace(/<c r="A\d+"[^/]*\/>|<c r="A\d+"[^>]*>[\s\S]*?<\/c>/, "");
        next = next.replace(/<c r="B\d+"[^/]*\/>|<c r="B\d+"[^>]*>[\s\S]*?<\/c>/, bXml);
        next = next.replace(/<c r="C\d+"[^/]*\/>|<c r="C\d+"[^>]*>[\s\S]*?<\/c>/, cXml);
        return next.replace(/(<row r="\d+"[^>]*>)/, "$1" + aXml);
      });
    }

    const extras = ["D", "E", "F", "G", "H", "I", "J", "K", "L"]
      .map(function (col) { return emptyCell(col + rowNum); })
      .join("");
    const row =
      '<row r="' + rowNum + '" spans="1:12" x14ac:dyDescent="0.25">' +
      aXml + bXml + cXml + extras + "</row>";
    return sheetXml.replace("</sheetData>", row + "</sheetData>");
  }

  function patchSheet(sheetXml, rowCells) {
    let xml = sheetXml;
    const rowNums = Object.keys(rowCells).map(Number).sort(function (a, b) { return a - b; });
    let maxRow = 63;

    rowNums.forEach(function (rowNum) {
      const cells = rowCells[rowNum];
      const aXml = textCell("A" + rowNum, cells.A);
      const bXml = cells.B == null ? emptyCell("B" + rowNum) : textCell("B" + rowNum, cells.B);
      const cXml = cells.C == null ? emptyCell("C" + rowNum) : textCell("C" + rowNum, cells.C);
      xml = replaceOrInsertRow(xml, rowNum, aXml, bXml, cXml);
      if (rowNum > maxRow) maxRow = rowNum;
    });

    return xml.replace(/dimension ref="A1:L\d+"/, 'dimension ref="A1:L' + maxRow + '"');
  }

  function buildWorkbook(rows) {
    if (!root.JSZip) {
      return Promise.reject(new Error("JSZip missing"));
    }
    if (!root.TEMPLATE_XLSX_BASE64) {
      return Promise.reject(new Error("template missing"));
    }

    return root.JSZip.loadAsync(base64ToUint8(root.TEMPLATE_XLSX_BASE64)).then(function (zip) {
      const sstFile = zip.file("xl/sharedStrings.xml");
      const sheetFile = zip.file("xl/worksheets/sheet1.xml");
      if (!sstFile || !sheetFile) {
        throw new Error("template structure invalid");
      }
      return Promise.all([sstFile.async("string"), sheetFile.async("string")]).then(function (parts) {
        const sstXml = parts[0];
        const countMatch = sstXml.match(/uniqueCount="(\d+)"/);
        const additions = [];
        const nextIndex = { value: countMatch ? Number(countMatch[1]) : 0 };
        const rowCells = {};

        rows.forEach(function (row, offset) {
          const excelRow = offset + 2;
          const phone = row.formatted || row.phone;
          rowCells[excelRow] = {
            A: addSharedString(additions, nextIndex, phone),
            B: row.firstName ? addSharedString(additions, nextIndex, row.firstName) : null,
            C: row.lastName ? addSharedString(additions, nextIndex, row.lastName) : null,
          };
        });

        zip.file("xl/sharedStrings.xml", patchSharedStrings(sstXml, additions, nextIndex.value));
        zip.file("xl/worksheets/sheet1.xml", patchSheet(parts[1], rowCells));
        return zip.generateAsync({
          type: "blob",
          mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          compression: "DEFLATE",
        });
      });
    });
  }

  root.TemplateExport = { buildWorkbook: buildWorkbook };
})(typeof window !== "undefined" ? window : globalThis);

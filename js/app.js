/* global XLSX, JsBarcode, JSZip, saveAs, ZXing, Html5Qrcode, Html5QrcodeSupportedFormats */

(function () {
  "use strict";

  // Tom & Jerry No. 103 — lembar kuning 12 label (3×4), kertas 20,5×16,5 cm, label 64×32 mm
  const LABEL_CONFIG = {
    SHEET_NAME: "Tom & Jerry No. 103 (12 label)",
    PAGE_WIDTH_MM: 205,
    PAGE_HEIGHT_MM: 165,
    COLS: 3,
    ROWS: 4,
    LABELS_PER_PAGE: 12,
    LABEL_WIDTH_MM: 64,
    LABEL_HEIGHT_MM: 32,
    DEFAULT_MARGIN_TOP_MM: 9,
    DEFAULT_COLUMN_LEFT_MM: [3, 66.5, 130],
    DEFAULT_VERTICAL_PITCH_MM: 38,
    FOOTER_HEIGHT_MM: 7.5,
    FOOTER_FONT_MM: 2,
    LOGO_SIZE_MM: 13,
    LOGO_ZONE_MM: 16,
    BARCODE_HEIGHT_MM: 9,
    BARCODE_MAX_WIDTH_MM: 38,
    PADDING_X_MM: 1.5,
    PADDING_Y_MM: 1,
    CODE_FONT_MM: 2.1,
    LOGO_GAP_MM: 0.5,
    CODE_GAP_MM: 0.35,
    RENDER_DPI: 300,
  };

  const EXCEL_CONFIG = {
    NAME_HEADERS: new Set(["nama", "name", "produk", "product", "nama produk", "product name"]),
    BARCODE_HEADERS: new Set(["kode barcode", "barcode", "kode", "code", "barcode code"]),
    DEFAULT_PREFIX: "BR",
    DEFAULT_PADDING: 3,
    BARCODE_COLUMN: "Kode Barcode",
    NAME_COLUMN: "Nama",
  };

  function mmToTwip(mm) {
    return Math.round(mm * 56.6929133858);
  }

  function mmToEmu(mm) {
    return Math.round((mm * 914400) / 25.4);
  }

  function mmToPx(mm) {
    return Math.round((mm / 25.4) * LABEL_CONFIG.RENDER_DPI);
  }

  function escapeXml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function formatBarcodeDisplay(code) {
    const normalized = String(code).trim().toUpperCase();
    const alphaNum = normalized.match(/^([A-Z]+)(\d+)$/);
    if (alphaNum) return alphaNum[1] + " - " + alphaNum[2];
    if (normalized.indexOf("-") >= 0) {
      return normalized.replace(/\s*-\s*/g, " - ");
    }
    return normalized.split("").join(" ");
  }

  function normalizeHeader(value) {
    if (value == null) return "";
    return String(value).trim().toLowerCase();
  }

  function findColumnIndex(headers, candidates) {
    for (let i = 0; i < headers.length; i += 1) {
      if (candidates.has(headers[i])) return i;
    }
    return -1;
  }

  function parseCodeConfig(prefixInput, startInput, digitsInput) {
    const prefix = String(prefixInput).trim().toUpperCase();
    const startNum = parseInt(String(startInput).trim(), 10);
    const padding = parseInt(String(digitsInput).trim(), 10);

    if (!prefix) {
      throw new Error("Awalan kode wajib diisi (contoh: SHM-ME-).");
    }
    if (!/^[A-Z0-9\-]+$/.test(prefix)) {
      throw new Error("Awalan kode hanya boleh huruf, angka, dan tanda minus (-).");
    }
    if (isNaN(startNum) || startNum < 0) {
      throw new Error("Nomor awal sequence tidak valid.");
    }
    if (isNaN(padding) || padding < 1 || padding > 10) {
      throw new Error("Digit sequence harus antara 1–10.");
    }

    return { prefix: prefix, startNum: startNum, padding: padding };
  }

  function buildCodeFromSequence(prefix, counter, padding) {
    return prefix + String(counter).padStart(padding, "0");
  }

  function updateCodePreview(prefixInput, startInput, digitsInput, previewEl) {
    try {
      const config = parseCodeConfig(prefixInput, startInput, digitsInput);
      const first = buildCodeFromSequence(config.prefix, config.startNum, config.padding);
      const second = buildCodeFromSequence(config.prefix, config.startNum + 1, config.padding);
      previewEl.textContent = first + ", " + second + ", ...";
    } catch (error) {
      previewEl.textContent = "-";
    }
  }

  function nextBarcodeCode(existingCodes, counter, prefix, padding) {
    while (true) {
      const code = buildCodeFromSequence(prefix, counter, padding);
      counter += 1;
      if (!existingCodes.has(code)) {
        existingCodes.add(code);
        return { code: code, counter: counter };
      }
    }
  }

  function sheetToRows(sheet) {
    if (!sheet || !sheet["!ref"]) return [];
    return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  }

  function rowsToSheet(rows) {
    return XLSX.utils.aoa_to_sheet(rows);
  }

  function createSampleExcel() {
    const rows = [
      [EXCEL_CONFIG.NAME_COLUMN],
      ["Sabun Lifebuoy"],
      ["Shampoo Clear 170ml"],
      ["Minyak Goreng 1L"],
      ["Teh Botol Sosro"],
      ["Air Mineral 600ml"],
    ];
    const sheet = rowsToSheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, "Produk");
    return XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  }

  function downloadSampleExcel() {
    const data = createSampleExcel();
    const blob = new Blob([data], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    saveAs(blob, "sample_produk.xlsx");
  }

  function processExcel(arrayBuffer, codeConfig) {
    let workbook;
    try {
      workbook = XLSX.read(arrayBuffer, { type: "array" });
    } catch (error) {
      throw new Error("File Excel tidak valid atau tidak bisa dibaca.");
    }

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = sheetToRows(sheet);

    if (rows.length === 0) throw new Error("File Excel kosong.");

    const headers = rows[0].map(normalizeHeader);
    const nameCol = findColumnIndex(headers, EXCEL_CONFIG.NAME_HEADERS);
    if (nameCol === -1) {
      throw new Error("Kolom nama produk tidak ditemukan. Pastikan ada kolom '" + EXCEL_CONFIG.NAME_COLUMN + "'.");
    }

    let barcodeCol = findColumnIndex(headers, EXCEL_CONFIG.BARCODE_HEADERS);
    if (barcodeCol === -1) {
      barcodeCol = headers.length;
      rows[0][barcodeCol] = EXCEL_CONFIG.BARCODE_COLUMN;
      headers.push(normalizeHeader(EXCEL_CONFIG.BARCODE_COLUMN));
    }

    const prefix = codeConfig.prefix;
    const padding = codeConfig.padding;
    const existingCodes = new Set();
    let counter = codeConfig.startNum;

    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const rawCode = rows[rowIndex][barcodeCol];
      if (rawCode != null && String(rawCode).trim()) {
        const code = String(rawCode).trim().toUpperCase();
        existingCodes.add(code);
        if (code.indexOf(prefix) === 0) {
          const suffix = code.slice(prefix.length);
          if (/^\d+$/.test(suffix)) {
            counter = Math.max(counter, Number(suffix) + 1);
          }
        }
      }
    }
    const products = [];

    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const nameValue = rows[rowIndex][nameCol];
      if (nameValue == null || !String(nameValue).trim()) continue;

      const name = String(nameValue).trim();
      const rawCode = rows[rowIndex][barcodeCol];
      let barcode;

      if (rawCode != null && String(rawCode).trim()) {
        barcode = String(rawCode).trim().toUpperCase();
      } else {
        const result = nextBarcodeCode(existingCodes, counter, prefix, padding);
        barcode = result.code;
        counter = result.counter;
        rows[rowIndex][barcodeCol] = barcode;
      }

      products.push({ name: name, barcode: barcode, rowIndex: rowIndex });
    }

    if (products.length === 0) {
      throw new Error("Tidak ada data produk yang valid di Excel.");
    }

    workbook.Sheets[sheetName] = rowsToSheet(rows);
    const updatedExcel = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    return { products: products, updatedExcel: updatedExcel };
  }

  function canvasToUint8Array(canvas) {
    const dataUrl = canvas.toDataURL("image/png");
    const base64 = dataUrl.split(",")[1];
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return bytes;
  }

  function normalizeHexColor(value, fallback) {
    const raw = String(value).trim();
    if (/^#[0-9A-Fa-f]{6}$/.test(raw)) return raw.toUpperCase();
    if (/^[0-9A-Fa-f]{6}$/.test(raw)) return ("#" + raw).toUpperCase();
    return fallback;
  }

  function parseFooterStyle(text, bgColor, textColor) {
    return {
      text: String(text).trim() || "PROPERTY OF PT SELARAS HASANAH MEDIKA",
      bgColor: normalizeHexColor(bgColor, "#000000"),
      textColor: normalizeHexColor(textColor, "#FFFFFF"),
    };
  }

  function syncColorInputs(colorInput, hexInput) {
    hexInput.value = colorInput.value.toUpperCase();
  }

  function syncHexToColor(colorInput, hexInput) {
    const normalized = normalizeHexColor(hexInput.value, colorInput.value);
    hexInput.value = normalized;
    colorInput.value = normalized;
  }

  function getPreviewCode() {
    try {
      const config = parseCodeConfig(
        codePrefixInput.value,
        codeStartInput.value,
        codeDigitsInput.value
      );
      return buildCodeFromSequence(config.prefix, config.startNum, config.padding);
    } catch (error) {
      return "SHM-ME-0001";
    }
  }

  function refreshLabelPreview() {
    const footerStyle = parseFooterStyle(
      footerTextInput.value,
      footerBgColorHexInput.value,
      footerTextColorHexInput.value
    );
    const code = getPreviewCode();
    const bytes = generateLabelPng(code, logoImage, footerStyle);
    const blob = new Blob([bytes], { type: "image/png" });
    if (labelPreviewObjectUrl) URL.revokeObjectURL(labelPreviewObjectUrl);
    labelPreviewObjectUrl = URL.createObjectURL(blob);
    labelPreviewImg.src = labelPreviewObjectUrl;
  }

  function drawFooter(ctx, text, width, footerTop, footerHeight, bgColor, textColor) {
    ctx.fillStyle = bgColor;
    ctx.fillRect(0, footerTop, width, footerHeight);

    ctx.fillStyle = textColor;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    let fontSize = mmToPx(LABEL_CONFIG.FOOTER_FONT_MM);
    ctx.font = "bold " + fontSize + "px Arial, sans-serif";

    while (fontSize > 5 && ctx.measureText(text).width > width - mmToPx(3)) {
      fontSize -= 1;
      ctx.font = "bold " + fontSize + "px Arial, sans-serif";
    }

    ctx.fillText(text, width / 2, footerTop + footerHeight / 2);
  }

  function drawBarcodeOnCanvas(ctx, code, x, y, width, height) {
    const tempCanvas = document.createElement("canvas");
    JsBarcode(tempCanvas, code, {
      format: "CODE128",
      displayValue: false,
      margin: 0,
      width: 2,
      height: 100,
    });
    ctx.drawImage(tempCanvas, x, y, width, height);
  }

  function generateLabelPng(code, logoImage, footerStyle) {
    const width = mmToPx(LABEL_CONFIG.LABEL_WIDTH_MM);
    const height = mmToPx(LABEL_CONFIG.LABEL_HEIGHT_MM);
    const footerHeight = mmToPx(LABEL_CONFIG.FOOTER_HEIGHT_MM);
    const topHeight = height - footerHeight;
    const padX = mmToPx(LABEL_CONFIG.PADDING_X_MM);
    const padY = mmToPx(LABEL_CONFIG.PADDING_Y_MM);
    const logoSize = logoImage ? mmToPx(LABEL_CONFIG.LOGO_SIZE_MM) : 0;
    const logoZoneWidth = logoImage ? mmToPx(LABEL_CONFIG.LOGO_ZONE_MM) : 0;
    const barcodeHeight = mmToPx(LABEL_CONFIG.BARCODE_HEIGHT_MM);
    const codeFontSize = mmToPx(LABEL_CONFIG.CODE_FONT_MM);
    const codeGap = mmToPx(LABEL_CONFIG.CODE_GAP_MM);

    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext("2d");

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);

    const contentLeft = logoZoneWidth > 0 ? logoZoneWidth : padX;
    const contentWidth = width - contentLeft - padX;

    const displayCode = formatBarcodeDisplay(code);
    ctx.font = "600 " + codeFontSize + "px Arial, sans-serif";
    const codeTextHeight = codeFontSize * 1.15;
    const barcodeMaxWidth = mmToPx(LABEL_CONFIG.BARCODE_MAX_WIDTH_MM);
    const barcodeWidth = Math.min(barcodeMaxWidth, contentWidth * 0.92);
    const blockHeight = barcodeHeight + codeGap + codeTextHeight;
    const blockTop = padY + Math.max(0, (topHeight - padY * 2 - blockHeight) / 2);
    const blockCenterY = blockTop + blockHeight / 2;
    const barcodeX = contentLeft + (contentWidth - barcodeWidth) / 2;

    if (logoImage) {
      const logoX = padX + Math.max(0, (logoZoneWidth - padX - logoSize) / 2);
      const logoY = blockCenterY - logoSize / 2;
      ctx.drawImage(logoImage, logoX, logoY, logoSize, logoSize);
    }

    drawBarcodeOnCanvas(ctx, code, barcodeX, blockTop, barcodeWidth, barcodeHeight);

    ctx.fillStyle = "#000000";
    ctx.textAlign = "center";
    ctx.textBaseline = "top";
    ctx.font = "600 " + codeFontSize + "px Arial, sans-serif";
    ctx.fillText(displayCode, contentLeft + contentWidth / 2, blockTop + barcodeHeight + codeGap);

    drawFooter(
      ctx,
      footerStyle.text,
      width,
      topHeight,
      footerHeight,
      footerStyle.bgColor,
      footerStyle.textColor
    );

    return canvasToUint8Array(canvas);
  }

  function generateLabelImages(products, logoImage, footerStyle) {
    const images = new Map();
    for (let i = 0; i < products.length; i += 1) {
      const product = products[i];
      if (!images.has(product.barcode)) {
        images.set(product.barcode, generateLabelPng(product.barcode, logoImage, footerStyle));
      }
    }
    return images;
  }

  function loadImageFromBytes(bytes) {
    return new Promise(function (resolve, reject) {
      const blob = new Blob([bytes], { type: "image/png" });
      const img = new Image();
      const url = URL.createObjectURL(blob);
      img.onload = function () {
        URL.revokeObjectURL(url);
        resolve(img);
      };
      img.onerror = function () {
        URL.revokeObjectURL(url);
        reject(new Error("Gagal memuat gambar label."));
      };
      img.src = url;
    });
  }

  async function buildLabelImageCache(labelImages) {
    const cache = new Map();
    const entries = Array.from(labelImages.entries());
    for (let i = 0; i < entries.length; i += 1) {
      cache.set(entries[i][0], await loadImageFromBytes(entries[i][1]));
    }
    return cache;
  }

  function getSheetLayout() {
    const marginTop = parseFloat(document.getElementById("calMarginTop").value);
    const col1 = parseFloat(document.getElementById("calCol1").value);
    const col2 = parseFloat(document.getElementById("calCol2").value);
    const col3 = parseFloat(document.getElementById("calCol3").value);
    const verticalPitch = parseFloat(document.getElementById("calPitchV").value);
    const defaultCols = LABEL_CONFIG.DEFAULT_COLUMN_LEFT_MM;

    return {
      pageWidthMm: LABEL_CONFIG.PAGE_WIDTH_MM,
      pageHeightMm: LABEL_CONFIG.PAGE_HEIGHT_MM,
      marginTopMm: isNaN(marginTop) ? LABEL_CONFIG.DEFAULT_MARGIN_TOP_MM : marginTop,
      columnLeftMm: [
        isNaN(col1) ? defaultCols[0] : col1,
        isNaN(col2) ? defaultCols[1] : col2,
        isNaN(col3) ? defaultCols[2] : col3,
      ],
      verticalPitchMm: isNaN(verticalPitch) ? LABEL_CONFIG.DEFAULT_VERTICAL_PITCH_MM : verticalPitch,
      labelWidthMm: LABEL_CONFIG.LABEL_WIDTH_MM,
      labelHeightMm: LABEL_CONFIG.LABEL_HEIGHT_MM,
      cols: LABEL_CONFIG.COLS,
      rows: LABEL_CONFIG.ROWS,
    };
  }

  async function generatePagePng(products, startIndex, labelImageCache, sheetLayout) {
    const pageWidth = mmToPx(sheetLayout.pageWidthMm);
    const pageHeight = mmToPx(sheetLayout.pageHeightMm);
    const labelWidth = mmToPx(sheetLayout.labelWidthMm);
    const labelHeight = mmToPx(sheetLayout.labelHeightMm);
    const marginTop = mmToPx(sheetLayout.marginTopMm);
    const verticalPitch = mmToPx(sheetLayout.verticalPitchMm);
    const columnLeft = sheetLayout.columnLeftMm.map(function (mm) {
      return mmToPx(mm);
    });

    const canvas = document.createElement("canvas");
    canvas.width = pageWidth;
    canvas.height = pageHeight;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, pageWidth, pageHeight);

    for (let rowIndex = 0; rowIndex < sheetLayout.rows; rowIndex += 1) {
      for (let colIndex = 0; colIndex < sheetLayout.cols; colIndex += 1) {
        const productIndex = startIndex + rowIndex * sheetLayout.cols + colIndex;
        if (productIndex >= products.length) continue;

        const product = products[productIndex];
        const labelImg = labelImageCache.get(product.barcode);
        if (!labelImg) continue;

        const x = columnLeft[colIndex];
        const y = marginTop + rowIndex * verticalPitch;
        ctx.drawImage(labelImg, x, y, labelWidth, labelHeight);
      }
    }

    return canvasToUint8Array(canvas);
  }

  function buildFullPageParagraph(relId, docPrId, pageBreakBefore) {
    const widthEmu = mmToEmu(LABEL_CONFIG.PAGE_WIDTH_MM);
    const heightEmu = mmToEmu(LABEL_CONFIG.PAGE_HEIGHT_MM);
    const pageBreakXml = pageBreakBefore ? "<w:pageBreakBefore/>" : "";

    return (
      "<w:p><w:pPr>" + pageBreakXml +
      "<w:spacing w:before=\"0\" w:after=\"0\" w:line=\"0\" w:lineRule=\"exact\"/>" +
      "<w:jc w:val=\"left\"/></w:pPr>" +
      "<w:r><w:drawing>" +
      "<wp:anchor distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" simplePos=\"0\" " +
      "relativeHeight=\"251658240\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\" " +
      "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">" +
      "<wp:simplePos x=\"0\" y=\"0\"/>" +
      "<wp:positionH relativeFrom=\"page\"><wp:posOffset>0</wp:posOffset></wp:positionH>" +
      "<wp:positionV relativeFrom=\"page\"><wp:posOffset>0</wp:posOffset></wp:positionV>" +
      "<wp:extent cx=\"" + widthEmu + "\" cy=\"" + heightEmu + "\"/>" +
      "<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>" +
      "<wp:wrapNone/>" +
      "<wp:docPr id=\"" + docPrId + "\" name=\"page\"/>" +
      "<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\" " +
      "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/></wp:cNvGraphicFramePr>" +
      "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
      "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
      "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
      "<pic:nvPicPr><pic:cNvPr id=\"0\" name=\"page\"/><pic:cNvPicPr/></pic:nvPicPr>" +
      "<pic:blipFill><a:blip r:embed=\"" + relId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
      "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
      "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" + widthEmu + "\" cy=\"" + heightEmu + "\"/></a:xfrm>" +
      "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>" +
      "</pic:pic></a:graphicData></a:graphic></wp:anchor>" +
      "</w:drawing></w:r></w:p>"
    );
  }

  async function generateLabelDocument(products, labelImages, sheetLayout) {
    const labelImageCache = await buildLabelImageCache(labelImages);
    const pageImages = [];
    let startIndex = 0;

    while (startIndex < products.length) {
      pageImages.push(await generatePagePng(products, startIndex, labelImageCache, sheetLayout));
      startIndex += LABEL_CONFIG.LABELS_PER_PAGE;
    }

    let bodyXml = "";
    for (let pageIndex = 0; pageIndex < pageImages.length; pageIndex += 1) {
      bodyXml += buildFullPageParagraph("rId" + (pageIndex + 2), pageIndex + 1, pageIndex > 0);
    }

    bodyXml +=
      "<w:sectPr>" +
      "<w:pgSz w:w=\"" + mmToTwip(LABEL_CONFIG.PAGE_WIDTH_MM) + "\" w:h=\"" + mmToTwip(LABEL_CONFIG.PAGE_HEIGHT_MM) + "\"/>" +
      "<w:pgMar w:top=\"0\" w:right=\"0\" w:bottom=\"0\" w:left=\"0\" " +
      "w:header=\"0\" w:footer=\"0\" w:gutter=\"0\"/>" +
      "</w:sectPr>";

    const documentXml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
      "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
      "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
      "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
      "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
      "<w:body>" + bodyXml + "</w:body></w:document>";

    const zip = new JSZip();
    zip.file(
      "[Content_Types].xml",
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
        "<Default Extension=\"png\" ContentType=\"image/png\"/>" +
        "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
        "</Types>"
    );

    zip.file(
      "_rels/.rels",
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
        "</Relationships>"
    );

    zip.file("word/document.xml", documentXml);

    let relsXml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";

    for (let pageIndex = 0; pageIndex < pageImages.length; pageIndex += 1) {
      const relId = "rId" + (pageIndex + 2);
      const fileName = "media/page" + (pageIndex + 1) + ".png";
      zip.file("word/" + fileName, pageImages[pageIndex]);
      relsXml +=
        "<Relationship Id=\"" + relId + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"" + fileName + "\"/>";
    }

    relsXml += "</Relationships>";
    zip.file("word/_rels/document.xml.rels", relsXml);

    return zip.generateAsync({ type: "blob" });
  }

  function loadImageFromFile(file) {
    return new Promise(function (resolve, reject) {
      const img = new Image();
      const url = URL.createObjectURL(file);
      img.onload = function () {
        resolve({ img: img, url: url });
      };
      img.onerror = function () {
        URL.revokeObjectURL(url);
        reject(new Error("Logo tidak bisa dibaca."));
      };
      img.src = url;
    });
  }

  const dropZone = document.getElementById("dropZone");
  const fileInput = document.getElementById("fileInput");
  const clearBtn = document.getElementById("clearBtn");
  const navItems = document.querySelectorAll(".nav-item");
  const tabPanels = document.querySelectorAll(".tab-panel");
  const sampleBtn = document.getElementById("sampleBtn");
  const generateBtn = document.getElementById("generateBtn");
  const fileInfo = document.getElementById("fileInfo");
  const fileName = document.getElementById("fileName");
  const resultSection = document.getElementById("resultSection");
  const resultMessage = document.getElementById("resultMessage");
  const downloadLink = document.getElementById("downloadLink");
  const loading = document.getElementById("loading");
  const toast = document.getElementById("toast");
  const footerTextInput = document.getElementById("footerText");
  const footerBgColorInput = document.getElementById("footerBgColor");
  const footerBgColorHexInput = document.getElementById("footerBgColorHex");
  const footerTextColorInput = document.getElementById("footerTextColor");
  const footerTextColorHexInput = document.getElementById("footerTextColorHex");
  const labelPreviewImg = document.getElementById("labelPreviewImg");
  const codePrefixInput = document.getElementById("codePrefix");
  const codeStartInput = document.getElementById("codeStart");
  const codeDigitsInput = document.getElementById("codeDigits");
  const codePreview = document.getElementById("codePreview");
  const logoInput = document.getElementById("logoInput");
  const logoBrowseBtn = document.getElementById("logoBrowseBtn");
  const logoClearBtn = document.getElementById("logoClearBtn");
  const appMain = document.querySelector(".app-main");
  const bottomBar = document.querySelector(".bottom-bar");
  const scanToggleBtn = document.getElementById("scanToggleBtn");
  const scanStatus = document.getElementById("scanStatus");
  const scanPhotoBtn = document.getElementById("scanPhotoBtn");
  const scanPhotoInput = document.getElementById("scanPhotoInput");
  const scanResultCard = document.getElementById("scanResultCard");
  const scanResultCode = document.getElementById("scanResultCode");
  const scanCopyBtn = document.getElementById("scanCopyBtn");
  const scanHistoryList = document.getElementById("scanHistoryList");
  const scanClearHistoryBtn = document.getElementById("scanClearHistoryBtn");

  let selectedFile = null;
  let downloadUrl = null;
  let logoImage = null;
  let logoObjectUrl = null;
  let labelPreviewObjectUrl = null;
  let isScanning = false;
  let scanHistoryItems = [];
  let lastScanCode = "";
  let lastScanTime = 0;
  let scanStream = null;
  let scanVideo = null;
  let scanLoopId = null;
  let zxingReader = null;
  let html5QrCode = null;
  let scanEngine = null;
  let scanStartPending = false;

  function switchTab(tabId) {
    if (tabId !== "tab-scan" && isScanning) {
      stopScanner();
    }

    tabPanels.forEach(function (panel) {
      panel.classList.toggle("active", panel.id === tabId);
    });
    navItems.forEach(function (item) {
      item.classList.toggle("active", item.getAttribute("data-tab") === tabId);
    });

    const hideGenerateBar = tabId === "tab-scan" || tabId === "tab-help";
    bottomBar.classList.toggle("hidden", hideGenerateBar);
    appMain.classList.toggle("no-bottom-bar", hideGenerateBar);

    if (tabId === "tab-scan" && !isScanning && !scanStartPending) {
      startLiveScanner();
    }
  }

  navItems.forEach(function (item) {
    item.addEventListener("click", function () {
      switchTab(item.getAttribute("data-tab"));
    });
  });

  function showToast(message, isError) {
    toast.textContent = message;
    toast.style.background = isError ? "#dc2626" : "#111827";
    toast.classList.remove("hidden");
    setTimeout(function () {
      toast.classList.add("hidden");
    }, 3500);
  }

  function setLoading(isLoading) {
    loading.classList.toggle("hidden", !isLoading);
    generateBtn.disabled = isLoading || !selectedFile;
  }

  function resetResult() {
    resultSection.classList.add("hidden");
    if (downloadUrl) {
      URL.revokeObjectURL(downloadUrl);
      downloadUrl = null;
    }
  }

  function setSelectedFile(file) {
    if (!file) {
      selectedFile = null;
      fileInfo.classList.add("hidden");
      generateBtn.disabled = true;
      resetResult();
      return;
    }

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      showToast("Gunakan file Excel (.xlsx).", true);
      return;
    }

    selectedFile = file;
    fileName.textContent = file.name;
    fileInfo.classList.remove("hidden");
    generateBtn.disabled = false;
    resetResult();
  }

  logoBrowseBtn.addEventListener("click", function () {
    logoInput.click();
  });

  function refreshCodePreview() {
    updateCodePreview(codePrefixInput.value, codeStartInput.value, codeDigitsInput.value, codePreview);
    refreshLabelPreview();
  }

  footerTextInput.addEventListener("input", refreshLabelPreview);

  footerBgColorInput.addEventListener("input", function () {
    syncColorInputs(footerBgColorInput, footerBgColorHexInput);
    refreshLabelPreview();
  });

  footerTextColorInput.addEventListener("input", function () {
    syncColorInputs(footerTextColorInput, footerTextColorHexInput);
    refreshLabelPreview();
  });

  footerBgColorHexInput.addEventListener("input", function () {
    syncHexToColor(footerBgColorInput, footerBgColorHexInput);
    refreshLabelPreview();
  });

  footerTextColorHexInput.addEventListener("input", function () {
    syncHexToColor(footerTextColorInput, footerTextColorHexInput);
    refreshLabelPreview();
  });

  codePrefixInput.addEventListener("input", refreshCodePreview);
  codeStartInput.addEventListener("input", refreshCodePreview);
  codeDigitsInput.addEventListener("input", refreshCodePreview);
  refreshCodePreview();

  logoInput.addEventListener("change", async function (event) {
    const file = event.target.files[0];
    if (!file) return;

    try {
      const loaded = await loadImageFromFile(file);
      if (logoObjectUrl) URL.revokeObjectURL(logoObjectUrl);
      logoImage = loaded.img;
      logoObjectUrl = loaded.url;
      logoClearBtn.classList.remove("hidden");
      refreshLabelPreview();
    } catch (error) {
      showToast(error.message || "Gagal load logo.", true);
    }
  });

  logoClearBtn.addEventListener("click", function () {
    logoImage = null;
    logoInput.value = "";
    if (logoObjectUrl) {
      URL.revokeObjectURL(logoObjectUrl);
      logoObjectUrl = null;
    }
    logoClearBtn.classList.add("hidden");
    refreshLabelPreview();
  });

  sampleBtn.addEventListener("click", function () {
    downloadSampleExcel();
  });

  fileInput.addEventListener("change", function (event) {
    setSelectedFile(event.target.files[0] || null);
  });

  clearBtn.addEventListener("click", function () {
    fileInput.value = "";
    setSelectedFile(null);
  });

  dropZone.addEventListener("click", function () {
    fileInput.click();
  });

  dropZone.addEventListener("dragover", function (event) {
    event.preventDefault();
    dropZone.classList.add("dragover");
  });

  dropZone.addEventListener("dragleave", function () {
    dropZone.classList.remove("dragover");
  });

  dropZone.addEventListener("drop", function (event) {
    event.preventDefault();
    dropZone.classList.remove("dragover");
    setSelectedFile(event.dataTransfer.files[0] || null);
  });

  generateBtn.addEventListener("click", async function () {
    if (!selectedFile) {
      showToast("Pilih file Excel terlebih dahulu.", true);
      return;
    }

    const footerStyle = parseFooterStyle(
      footerTextInput.value,
      footerBgColorHexInput.value,
      footerTextColorHexInput.value
    );
    let codeConfig;
    try {
      codeConfig = parseCodeConfig(codePrefixInput.value, codeStartInput.value, codeDigitsInput.value);
    } catch (error) {
      showToast(error.message, true);
      return;
    }

    setLoading(true);

    try {
      const arrayBuffer = await selectedFile.arrayBuffer();
      const result = processExcel(arrayBuffer, codeConfig);
      const labelImages = generateLabelImages(result.products, logoImage, footerStyle);
      const sheetLayout = getSheetLayout();
      const docxBlob = await generateLabelDocument(result.products, labelImages, sheetLayout);

      const zip = new JSZip();
      zip.file("labels.docx", docxBlob);
      zip.file("products_updated.xlsx", result.updatedExcel);
      const zipBlob = await zip.generateAsync({ type: "blob" });

      if (downloadUrl) URL.revokeObjectURL(downloadUrl);
      downloadUrl = URL.createObjectURL(zipBlob);
      downloadLink.href = downloadUrl;
      resultMessage.textContent = "Berhasil generate " + result.products.length + " label barcode.";
      resultSection.classList.remove("hidden");
      switchTab("tab-upload");
      resultSection.scrollIntoView({ behavior: "smooth", block: "nearest" });
      showToast("Barcode berhasil dibuat!", false);
    } catch (error) {
      showToast(error.message || "Terjadi kesalahan.", true);
    } finally {
      setLoading(false);
    }
  });

  function renderScanHistory() {
    scanHistoryList.innerHTML = "";

    if (scanHistoryItems.length === 0) {
      const emptyItem = document.createElement("li");
      emptyItem.className = "scan-history-empty";
      emptyItem.textContent = "Belum ada scan";
      scanHistoryList.appendChild(emptyItem);
      scanClearHistoryBtn.classList.add("hidden");
      return;
    }

    scanClearHistoryBtn.classList.remove("hidden");
    scanHistoryItems.forEach(function (code) {
      const item = document.createElement("li");
      const label = document.createElement("span");
      label.textContent = code;
      const copyBtn = document.createElement("button");
      copyBtn.type = "button";
      copyBtn.textContent = "Salin";
      copyBtn.addEventListener("click", function () {
        copyScanCode(code);
      });
      item.appendChild(label);
      item.appendChild(copyBtn);
      scanHistoryList.appendChild(item);
    });
  }

  function copyScanCode(code) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(code).then(function () {
        showToast("Kode disalin: " + code, false);
      }).catch(function () {
        showToast("Gagal salin kode.", true);
      });
      return;
    }

    showToast("Salin manual: " + code, false);
  }

  function handleScanSuccess(decodedText) {
    const code = String(decodedText).trim().toUpperCase();
    if (!code) return;

    const now = Date.now();
    if (code === lastScanCode && now - lastScanTime < 2000) return;

    lastScanCode = code;
    lastScanTime = now;
    scanResultCode.textContent = code;
    scanResultCard.classList.remove("hidden");

    if (scanHistoryItems[0] !== code) {
      scanHistoryItems.unshift(code);
      if (scanHistoryItems.length > 30) scanHistoryItems.pop();
      renderScanHistory();
    }

    if (navigator.vibrate) navigator.vibrate(80);
    showToast("Terbaca: " + code, false);
  }

  function setScanButtonState(active) {
    scanToggleBtn.classList.remove("hidden");
    if (active) {
      scanToggleBtn.textContent = "⏹ Stop Kamera";
      scanToggleBtn.classList.remove("btn-primary");
      scanToggleBtn.classList.add("btn-danger");
      scanStatus.textContent = "Arahkan ke barcode Code128...";
      return;
    }

    scanToggleBtn.textContent = "🔄 Coba Lagi";
    scanToggleBtn.classList.remove("btn-danger");
    scanToggleBtn.classList.add("btn-primary");
  }

  function clearScanViewport() {
    scanViewport.innerHTML = "";
  }

  function showScanPlaceholder() {
    scanViewport.innerHTML = '<div class="scan-placeholder">📷 Kamera live</div>';
  }

  async function stopHtml5Scanner() {
    if (!html5QrCode) return;

    try {
      await html5QrCode.stop();
    } catch (error) {
      // ignore
    }

    try {
      html5QrCode.clear();
    } catch (error) {
      // ignore
    }

    html5QrCode = null;
  }

  async function pickBackCameraId() {
    if (typeof Html5Qrcode === "undefined" || !Html5Qrcode.getCameras) {
      return { facingMode: "environment" };
    }

    try {
      const cameras = await Html5Qrcode.getCameras();
      if (!cameras || !cameras.length) {
        return { facingMode: "environment" };
      }

      const backCamera = cameras.find(function (camera) {
        return /back|rear|environment|belakang/i.test(camera.label);
      });

      if (backCamera) return backCamera.id;
      if (cameras.length > 1) return cameras[cameras.length - 1].id;
      return cameras[0].id;
    } catch (error) {
      return { facingMode: "environment" };
    }
  }

  async function startHtml5LiveScanner() {
    if (typeof Html5Qrcode === "undefined") {
      throw new Error("Html5Qrcode tidak termuat");
    }

    clearScanViewport();
    html5QrCode = new Html5Qrcode("scanViewport", { verbose: false });
    const cameraId = await pickBackCameraId();

    const config = {
      fps: 12,
      qrbox: function (viewfinderWidth, viewfinderHeight) {
        return {
          width: Math.min(Math.floor(viewfinderWidth * 0.92), 340),
          height: Math.min(Math.floor(viewfinderHeight * 0.4), 150),
        };
      },
      formatsToSupport: [Html5QrcodeSupportedFormats.CODE_128],
      experimentalFeatures: {
        useBarCodeDetectorIfSupported: true,
      },
      videoConstraints: {
        facingMode: "environment",
        width: { min: 640, ideal: 1280 },
        height: { min: 480, ideal: 720 },
      },
    };

    await html5QrCode.start(cameraId, config, handleScanSuccess, function () {});
    scanEngine = "html5";
  }

  function stopScanStream() {
    if (scanLoopId) {
      clearInterval(scanLoopId);
      scanLoopId = null;
    }

    if (scanStream) {
      scanStream.getTracks().forEach(function (track) {
        track.stop();
      });
      scanStream = null;
    }

    scanVideo = null;
  }

  function stopZxingReader() {
    if (zxingReader) {
      try {
        zxingReader.reset();
      } catch (error) {
        // ignore
      }
      zxingReader = null;
    }
  }

  function isSecureScanContext() {
    return window.isSecureContext === true;
  }

  async function getCameraStream() {
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      throw new Error("Browser tidak mendukung kamera.");
    }

    const attempts = [
      {
        audio: false,
        video: {
          facingMode: { ideal: "environment" },
          width: { ideal: 1280 },
          height: { ideal: 720 },
        },
      },
      { audio: false, video: { facingMode: { ideal: "environment" } } },
      { audio: false, video: { facingMode: "environment" } },
      { audio: false, video: true },
    ];

    let lastError = null;
    for (let i = 0; i < attempts.length; i += 1) {
      try {
        return await navigator.mediaDevices.getUserMedia(attempts[i]);
      } catch (error) {
        lastError = error;
      }
    }

    throw lastError || new Error("Gagal akses kamera");
  }

  function createScanVideoElement(stream) {
    const video = document.createElement("video");
    video.setAttribute("playsinline", "true");
    video.setAttribute("webkit-playsinline", "true");
    video.setAttribute("autoplay", "true");
    video.muted = true;
    video.playsInline = true;
    video.srcObject = stream;
    return video;
  }

  async function playScanVideo(video) {
    return new Promise(function (resolve, reject) {
      video.onloadedmetadata = function () {
        video
          .play()
          .then(resolve)
          .catch(reject);
      };
      video.onerror = function () {
        reject(new Error("Video kamera gagal dimuat"));
      };
    });
  }

  function startDetectorLoop(video) {
    const detector = new BarcodeDetector({ formats: ["code_128"] });
    scanLoopId = setInterval(function () {
      if (!isScanning || !video || video.readyState < 2) return;
      detector
        .detect(video)
        .then(function (codes) {
          if (codes && codes.length > 0) {
            handleScanSuccess(codes[0].rawValue);
          }
        })
        .catch(function () {});
    }, 300);
  }

  function startZxingVideoLoop(video) {
    const hints = new Map();
    hints.set(ZXing.DecodeHintType.POSSIBLE_FORMATS, [ZXing.BarcodeFormat.CODE_128]);
    hints.set(ZXing.DecodeHintType.TRY_HARDER, true);

    zxingReader = new ZXing.BrowserMultiFormatReader(hints, 350);

    if (typeof zxingReader.decodeFromVideoElementContinuously === "function") {
      zxingReader.decodeFromVideoElementContinuously(video, function (result) {
        if (result) handleScanSuccess(result.getText());
      });
      return;
    }

    scanLoopId = setInterval(function () {
      if (!isScanning || !video || video.readyState < 2 || !zxingReader) return;
      zxingReader
        .decodeFromVideoElement(video)
        .then(function (result) {
          handleScanSuccess(result.getText());
        })
        .catch(function () {});
    }, 400);
  }

  function getScanErrorMessage(error) {
    const name = error && error.name ? error.name : "";
    if (name === "NotAllowedError" || name === "PermissionDeniedError") {
      return "Kamera ditolak. Pakai Ambil Foto Barcode, atau allow kamera di setting browser.";
    }
    if (name === "NotFoundError" || name === "DevicesNotFoundError") {
      return "Kamera tidak ditemukan di perangkat ini.";
    }
    if (name === "NotReadableError" || name === "TrackStartError") {
      return "Kamera sedang dipakai app lain. Tutup app lain lalu coba lagi.";
    }
    if (name === "SecurityError" || !isSecureScanContext()) {
      return "Buka lewat HTTPS. Kamera tidak jalan di file lokal.";
    }
    return "Gagal buka kamera. Coba tombol Ambil Foto Barcode.";
  }

  async function startLiveScanner() {
    if (isScanning || scanStartPending) return;

    if (!isSecureScanContext()) {
      scanStatus.textContent = "Butuh HTTPS agar kamera live jalan";
      showToast("Buka lewat HTTPS (GitHub Pages).", true);
      return;
    }

    scanStartPending = true;
    scanStatus.textContent = "Meminta izin kamera...";
    stopScanStream();
    stopZxingReader();
    await stopHtml5Scanner();

    try {
      if (typeof Html5Qrcode !== "undefined") {
        await startHtml5LiveScanner();
        isScanning = true;
        setScanButtonState(true);
        scanStartPending = false;
        return;
      }
    } catch (error) {
      await stopHtml5Scanner();
    }

    try {
      scanStream = await getCameraStream();
      clearScanViewport();
      scanVideo = createScanVideoElement(scanStream);
      scanViewport.appendChild(scanVideo);
      await playScanVideo(scanVideo);

      if ("BarcodeDetector" in window) {
        scanEngine = "detector";
        startDetectorLoop(scanVideo);
      } else if (typeof ZXing !== "undefined") {
        scanEngine = "zxing";
        startZxingVideoLoop(scanVideo);
      } else {
        throw new Error("Scanner tidak tersedia");
      }

      isScanning = true;
      setScanButtonState(true);
    } catch (error) {
      stopScanStream();
      stopZxingReader();
      showScanPlaceholder();
      isScanning = false;
      setScanButtonState(false);
      scanStatus.textContent = getScanErrorMessage(error);
      showToast(getScanErrorMessage(error), true);
    }

    scanStartPending = false;
  }

  async function stopScanner() {
    stopScanStream();
    stopZxingReader();
    await stopHtml5Scanner();
    showScanPlaceholder();
    isScanning = false;
    scanEngine = null;
    scanStartPending = false;
    setScanButtonState(false);
  }

  function decodeScanFromImage(file) {
    decodeScanFromImageFile(file);
  }

  async function decodeScanFromImageFile(file) {
    if (!file) return;

    scanStatus.textContent = "Membaca foto...";

    const url = URL.createObjectURL(file);
    const img = new Image();

    try {
      await new Promise(function (resolve, reject) {
        img.onload = resolve;
        img.onerror = reject;
        img.src = url;
      });

      if ("BarcodeDetector" in window) {
        try {
          const detector = new BarcodeDetector({ formats: ["code_128"] });
          const codes = await detector.detect(img);
          if (codes && codes.length > 0) {
            handleScanSuccess(codes[0].rawValue);
            return;
          }
        } catch (error) {
          // fallback ke ZXing
        }
      }

      if (typeof ZXing === "undefined" || !ZXing.BrowserMultiFormatReader) {
        throw new Error("Library scanner belum termuat");
      }

      const hints = new Map();
      hints.set(ZXing.DecodeHintType.POSSIBLE_FORMATS, [ZXing.BarcodeFormat.CODE_128]);
      hints.set(ZXing.DecodeHintType.TRY_HARDER, true);
      const reader = new ZXing.BrowserMultiFormatReader(hints);
      const result = await reader.decodeFromImageElement(img);
      handleScanSuccess(result.getText());
    } catch (error) {
      showToast("Barcode tidak terbaca. Foto lebih dekat, fokus, & pencahayaan cukup.", true);
      scanStatus.textContent = isScanning
        ? "Arahkan ke barcode Code128..."
        : "Tap tab Scan untuk buka kamera";
    } finally {
      URL.revokeObjectURL(url);
    }
  }

  scanToggleBtn.addEventListener("click", function () {
    if (isScanning) stopScanner();
    else startLiveScanner();
  });

  scanPhotoBtn.addEventListener("click", function () {
    scanPhotoInput.click();
  });

  scanPhotoInput.addEventListener("change", function (event) {
    decodeScanFromImage(event.target.files[0] || null);
    scanPhotoInput.value = "";
  });

  scanCopyBtn.addEventListener("click", function () {
    copyScanCode(scanResultCode.textContent || "");
  });

  scanClearHistoryBtn.addEventListener("click", function () {
    scanHistoryItems = [];
    lastScanCode = "";
    lastScanTime = 0;
    renderScanHistory();
    showToast("Riwayat scan dihapus.", false);
  });

  renderScanHistory();
})();

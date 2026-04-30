const MAX_FILE_SIZE = 5 * 1024 * 1024; // 5MB
const SUPPORTED_EXTENSIONS = [".xlsx", ".xls"];
const state = {
  workbookData: [],
  cleanedData: [],
  cleanSummary: null,
  fileName: "",
};
const fileInput = document.getElementById("fileInput");
const cleanButton = document.getElementById("cleanButton");
const downloadButton = document.getElementById("downloadButton");
const backupConfirm = document.getElementById("backupConfirm");
const statusText = document.getElementById("statusText");
const loadingIndicator = document.getElementById("loadingIndicator");
const previewThead = document.querySelector("#previewTable thead");
const previewTbody = document.querySelector("#previewTable tbody");
const setStatus = (text, isError = false) => {
  statusText.textContent = text;
  statusText.className = isError ? "text-sm text-red-600" : "text-sm text-slate-600";
};
const setLoading = (loading) => {
  loadingIndicator.classList.toggle("hidden", !loading);
};
const clearPreview = () => {
  previewThead.innerHTML = "";
  previewTbody.innerHTML = "";
};
const updateDownloadButtonState = () => {
  downloadButton.disabled = !state.cleanedData.length || !backupConfirm.checked;
};
const normalizeHeaderBase = (headerText) =>
  String(headerText || "")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/[#\$%&\*@!^~`+=\[\]{}()|\\;:'",.<>/?-]/g, "")
    .replace(/\s+/g, " ");
const toCamelCase = (input) => {
  const words = input
    .toLowerCase()
    .split(" ")
    .filter(Boolean);
  return words
    .map((word, index) =>
      index === 0 ? word : `${word.charAt(0).toUpperCase()}${word.slice(1)}`
    )
    .join("");
};
const toSnakeCase = (input) =>
  input
    .toLowerCase()
    .split(" ")
    .filter(Boolean)
    .join("_");
const sanitizeCellValue = (value) => {
  if (value === undefined || value === null) {
    return "";
  }
  return String(value)
    .normalize("NFKC")
    .trim()
    .replace(/[^\u3400-\u9FFF\uF900-\uFAFFA-Za-z0-9_ ]+/g, "");
};
const makeHeadersUnique = (headers) => {
  const counts = new Map();
  return headers.map((header, index) => {
    const safeHeader = header || `column_${index + 1}`;
    const currentCount = counts.get(safeHeader) || 0;
    counts.set(safeHeader, currentCount + 1);
    return currentCount === 0 ? safeHeader : `${safeHeader}_${currentCount + 1}`;
  });
};
const cleanHeaders = (headers, style) => {
  const cleaned = headers.map((header, index) => {
    const base = normalizeHeaderBase(header);
    if (!base) {
      return `column_${index + 1}`;
    }
    if (style === "camel") {
      return toCamelCase(base);
    }
    if (style === "snake") {
      return toSnakeCase(base);
    }
    // 默认模式：移除空格与特殊字符
    return base.replace(/\s+/g, "");
  });
  return makeHeadersUnique(cleaned);
};
const getSelectedStyle = () =>
  document.querySelector('input[name="namingStyle"]:checked')?.value || "default";
const renderPreview = (rows) => {
  clearPreview();
  if (!rows.length) {
    return;
  }
  const headers = Object.keys(rows[0]);
  const headRow = document.createElement("tr");
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headRow.appendChild(th);
  });
  previewThead.appendChild(headRow);
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    headers.forEach((header) => {
      const td = document.createElement("td");
      const value = row[header];
      td.textContent = value === undefined || value === null ? "" : String(value);
      tr.appendChild(td);
    });
    previewTbody.appendChild(tr);
  });
};
const parseWorkbook = async (file) => {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("Excel 文件中未找到工作表。");
  }
  const worksheet = workbook.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, {
    defval: "",
    raw: false,
  });
  return rows;
};
const exportToCsv = (rows, fileName) => {
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const csvContent = XLSX.utils.sheet_to_csv(worksheet);
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `${fileName || "cleaned_data"}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
};
const cleanDataHeaders = () => {
  if (!state.workbookData.length) {
    throw new Error("当前没有可处理的数据。");
  }
  const style = getSelectedStyle();
  const originalHeaders = Object.keys(state.workbookData[0]);
  const newHeaders = cleanHeaders(originalHeaders, style);
  // 按列索引映射旧列名到新列名，确保所有行都被完整转换
  const headerMap = originalHeaders.map((oldHeader, index) => ({
    oldHeader,
    newHeader: newHeaders[index],
  }));
  const cleanedRows = state.workbookData.map((row) => {
    const nextRow = {};
    headerMap.forEach(({ oldHeader, newHeader }) => {
      nextRow[newHeader] = sanitizeCellValue(row[oldHeader]);
    });
    return nextRow;
  });
  const uniqueRows = [];
  const seen = new Set();
  cleanedRows.forEach((row) => {
    const rowKey = newHeaders.map((header) => row[header]).join("||");
    if (seen.has(rowKey)) {
      return;
    }
    seen.add(rowKey);
    uniqueRows.push(row);
  });
  return {
    rows: uniqueRows,
    originalRowCount: cleanedRows.length,
    uniqueRowCount: uniqueRows.length,
    removedDuplicates: cleanedRows.length - uniqueRows.length,
  };
};
fileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  state.workbookData = [];
  state.cleanedData = [];
  state.cleanSummary = null;
  backupConfirm.checked = false;
  updateDownloadButtonState();
  cleanButton.disabled = true;
  clearPreview();
  if (!file) {
    setStatus("未选择文件。");
    return;
  }
  const lowerName = file.name.toLowerCase();
  const isSupported = SUPPORTED_EXTENSIONS.some((ext) => lowerName.endsWith(ext));
  if (!isSupported) {
    setStatus("不支持的文件类型，请上传 .xlsx 或 .xls 文件。", true);
    return;
  }
  if (file.size > MAX_FILE_SIZE) {
    setStatus("文件超过 5MB，请选择更小的文件。", true);
    return;
  }
  setLoading(true);
  setStatus("正在读取并解析 Excel 文件...");
  try {
    const rows = await parseWorkbook(file);
    if (!rows.length) {
      throw new Error("工作表为空，无法预览或导出。");
    }
    state.workbookData = rows;
    state.fileName = file.name.replace(/\.(xlsx|xls)$/i, "");
    cleanButton.disabled = false;
    renderPreview(rows.slice(0, 5));
    updateDownloadButtonState();
    setStatus(`文件加载成功，共 ${rows.length} 行。请点击“开始清洗列名”。`);
  } catch (error) {
    setStatus(`读取失败：${error.message}`, true);
  } finally {
    setLoading(false);
  }
});
cleanButton.addEventListener("click", async () => {
  if (!state.workbookData.length) {
    setStatus("请先上传并读取文件。", true);
    return;
  }
  setLoading(true);
  setStatus("正在清洗列名...");
  try {
    // 使用 Promise.resolve 包裹，统一 async/await 风格
    const { rows, originalRowCount, uniqueRowCount, removedDuplicates } = await Promise.resolve(
      cleanDataHeaders()
    );
    state.cleanedData = rows;
    state.cleanSummary = {
      originalRowCount,
      uniqueRowCount,
      removedDuplicates,
    };
    renderPreview(state.cleanedData.slice(0, 5));
    updateDownloadButtonState();
    setStatus(`清洗完成，移除了 ${removedDuplicates} 个重复行。`);
  } catch (error) {
    setStatus(`清洗失败：${error.message}`, true);
  } finally {
    setLoading(false);
  }
});
backupConfirm.addEventListener("change", () => {
  updateDownloadButtonState();
  if (!backupConfirm.checked && state.cleanedData.length) {
    setStatus("请先确认已备份原始文件，再下载 CSV。", true);
    return;
  }
  if (backupConfirm.checked && state.cleanedData.length) {
    setStatus("已确认备份。请复核预览数据后下载 CSV。");
  }
});
downloadButton.addEventListener("click", () => {
  if (!state.cleanedData.length) {
    setStatus("没有可下载的数据，请先完成清洗。", true);
    return;
  }
  if (!backupConfirm.checked) {
    setStatus("请先勾选“已确认备份”后再下载。", true);
    updateDownloadButtonState();
    return;
  }
  try {
    const summary = state.cleanSummary || {
      originalRowCount: state.workbookData.length,
      uniqueRowCount: state.cleanedData.length,
      removedDuplicates: Math.max(state.workbookData.length - state.cleanedData.length, 0),
    };
    const confirmed = window.confirm(
      `清洗完成！
共处理了 ${summary.originalRowCount} 行数据，清洗后剩余 ${summary.uniqueRowCount} 行（移除了 ${summary.removedDuplicates} 个重复行）。
请确认预览数据无误。
点击“确定”开始下载CSV文件，点击“取消”返回。`
    );
    if (!confirmed) {
      return;
    }
    exportToCsv(state.cleanedData, `${state.fileName}_cleaned`);
    setStatus("CSV 下载已开始。");
  } catch (error) {
    setStatus(`下载失败：${error.message}`, true);
  }
});
// 页面关闭时自动释放内存中的数据引用
window.addEventListener("beforeunload", () => {
  state.workbookData = [];
  state.cleanedData = [];
  state.cleanSummary = null;
  state.fileName = "";
});
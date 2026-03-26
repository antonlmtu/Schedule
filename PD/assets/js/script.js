document.addEventListener("DOMContentLoaded", function () {
  const downloadPdfBtn = document.getElementById("downloadPdfBtn");
  const copyImageBtn = document.getElementById("copyImageBtn");
  const captureArea = document.getElementById("captureArea");

  const titleLine = document.getElementById("titleLine");
  const editableSuffix = document.getElementById("editableSuffix");

  const scheduleTableBody = document.getElementById("scheduleTableBody");

  const memberListInput = document.getElementById("memberListInput");
  const memberListFile = document.getElementById("memberListFile");
  const memberListCount = document.getElementById("memberListCount");
  const memberListStatus = document.getElementById("memberListStatus");
  const randomizeAssignmentsBtn = document.getElementById(
    "randomizeAssignmentsBtn"
  );
  const clearListBtn = document.getElementById("clearListBtn");

  const dayColumns = [4, 5, 6, 7];
  const sourceRows = [2, 3, 4, 5, 6, 7, 8];
  const primaryManualRows = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11];
  const generatedEditableRows = [
    12, 13, 14, 15, 16, 17, 18, 19,
    20, 21, 22, 23, 24, 25, 26, 27,
    28, 29, 30, 31, 32, 33, 34,
    35, 36, 37, 38, 39, 40, 41, 42, 43, 44,
  ];
  const freeInputRows = [45];
  const manualRows = [...primaryManualRows, ...generatedEditableRows, ...freeInputRows];
  const validatedRows = [...primaryManualRows, ...generatedEditableRows];
  const permuteRows = [28, 29, 30, 31, 32, 33, 34];
  const randomPoolRows = [
    12, 13, 14, 15, 16, 17, 18, 19,
    20, 21, 22, 23, 24, 25, 26, 27,
    35, 36, 37, 38, 39, 40, 41, 42, 43, 44,
  ];
  const generatedRows = [...permuteRows, ...randomPoolRows];

  const cellRegistry = new Map();
  const cellFlashTimers = new Map();
  let parsedMemberList = [];
  const columnRefreshTimers = new Map();
  let listRefreshTimer = null;
  let activeColumn = dayColumns[0] || null;

  function sanitizeDisplayText(value = "", { trim = true } = {}) {
    const normalized = value
      .replace(/[\r\n]+/g, " ")
      .replace(/\u00A0/g, " ")
      .replace(/[ ]{2,}/g, " ");

    return trim ? normalized.trim() : normalized;
  }

  function normalizeMemberKey(value = "") {
    return sanitizeDisplayText(value, { trim: true }).toUpperCase();
  }

  function formatVietnameseNameCase(value = "") {
    const cleaned = sanitizeDisplayText(value, { trim: false }).toLocaleLowerCase("vi-VN");
    let result = "";
    let shouldCapitalize = true;

    for (const char of cleaned) {
      if (/\p{L}/u.test(char)) {
        result += shouldCapitalize ? char.toLocaleUpperCase("vi-VN") : char;
        shouldCapitalize = false;
      } else {
        result += char;

        if (char === ".") {
          shouldCapitalize = true;
        }
      }
    }

    return result.replace(/\.\s*/g, ". ");
  }

  function normalizeLooseMemberKey(value = "") {
    return sanitizeDisplayText(value, { trim: true })
      .normalize("NFD")
      .replace(/[̀-ͯ]/g, "")
      .replace(/đ/g, "d")
      .replace(/Đ/g, "D")
      .toUpperCase();
  }

  function getCanonicalMemberName(value = "") {
    const displayName = sanitizeDisplayText(value, { trim: true });
    if (!displayName) return "";

    const exactKey = normalizeMemberKey(displayName);
    const looseKey = normalizeLooseMemberKey(displayName);

    return (
      parsedMemberList.find((name) => normalizeMemberKey(name) === exactKey) ||
      parsedMemberList.find((name) => normalizeLooseMemberKey(name) === looseKey) ||
      ""
    );
  }

  function rememberCellValidValue(cell, value = "") {
    if (!cell) return;

    const safeValue = sanitizeDisplayText(value, { trim: true });

    if (safeValue) {
      cell.dataset.lastValidValue = safeValue;
      return;
    }

    delete cell.dataset.lastValidValue;
  }

  function clearValidationState(cell) {
    if (!cell) return;
    cell.classList.remove("cell-invalid", "cell-auto-replaced");
  }

  function flashCellState(cell, className) {
    if (!cell || !className) return;

    const cellKey = `${cell.dataset.row || "x"}-${cell.dataset.col || "x"}-${className}`;
    const existingTimer = cellFlashTimers.get(cellKey);

    if (existingTimer) {
      window.clearTimeout(existingTimer);
    }

    cell.classList.add(className);

    cellFlashTimers.set(
      cellKey,
      window.setTimeout(() => {
        cell.classList.remove(className);
        cellFlashTimers.delete(cellKey);
      }, 1800)
    );
  }

  function registerCell(row, col, cell) {
    if (!cell || !row || !col) return;
    cell.dataset.row = String(row);
    cell.dataset.col = String(col);
    cellRegistry.set(`${row}-${col}`, cell);
  }

  function getCell(row, col) {
    return cellRegistry.get(`${row}-${col}`) || null;
  }

  function getCellDisplayValue(row, col) {
    const cell = getCell(row, col);
    if (!cell) return "";
    return sanitizeDisplayText(cell.textContent || "", { trim: true });
  }

  function getCellValue(row, col) {
    return normalizeMemberKey(getCellDisplayValue(row, col));
  }

  function setCellValue(row, col, value = "", { rememberValid = false } = {}) {
    const cell = getCell(row, col);
    if (!cell) return;

    const safeValue = sanitizeDisplayText(value, { trim: true });
    cell.textContent = safeValue;

    if (rememberValid) {
      rememberCellValidValue(cell, safeValue);
    }
  }

  function autoReplaceDuplicateSourceMembers(col, preferredRow = null) {
    if (!dayColumns.includes(col) || !parsedMemberList.length) {
      return {
        replacedRows: [],
        clearedRows: [],
      };
    }

    const groupedRows = new Map();
    const orderedEntries = sourceRows
      .map((row) => {
        const displayName = getCellDisplayValue(row, col);
        const canonicalName = getCanonicalMemberName(displayName);

        if (!canonicalName) return null;

        return {
          row,
          canonicalName,
          key: normalizeMemberKey(canonicalName),
        };
      })
      .filter(Boolean);

    orderedEntries.forEach((entry) => {
      if (!groupedRows.has(entry.key)) {
        groupedRows.set(entry.key, []);
      }

      groupedRows.get(entry.key).push(entry);
    });

    const rowsToKeep = new Set();
    const rowsToReplace = [];

    groupedRows.forEach((entries) => {
      const keepEntry =
        (preferredRow && entries.find((entry) => entry.row === preferredRow)) ||
        entries[0];

      rowsToKeep.add(keepEntry.row);

      entries.forEach((entry) => {
        if (entry.row !== keepEntry.row) {
          rowsToReplace.push(entry.row);
        }
      });
    });

    const usedKeys = new Set(
      orderedEntries
        .filter((entry) => rowsToKeep.has(entry.row))
        .map((entry) => entry.key)
    );

    const availableMembers = parsedMemberList.filter(
      (name) => !usedKeys.has(normalizeMemberKey(name))
    );

    const summary = {
      replacedRows: [],
      clearedRows: [],
    };

    rowsToReplace.sort((firstRow, secondRow) => firstRow - secondRow);

    rowsToReplace.forEach((row) => {
      const replacementName = availableMembers.shift() || "";
      const cell = getCell(row, col);

      setCellValue(row, col, replacementName, {
        rememberValid: Boolean(replacementName),
      });
      clearValidationState(cell);

      if (replacementName) {
        usedKeys.add(normalizeMemberKey(replacementName));
        summary.replacedRows.push({ row, replacementName });
        flashCellState(cell, "cell-auto-replaced");
        return;
      }

      rememberCellValidValue(cell, "");
      summary.clearedRows.push({ row });
      flashCellState(cell, "cell-invalid");
    });

    return summary;
  }

  function validateManualEntryCell(cell, { announce = true } = {}) {
    const row = Number(cell?.dataset?.row || 0);
    const col = Number(cell?.dataset?.col || 0);
    const summary = {
      isValid: true,
      announced: false,
      replacedRows: [],
      clearedRows: [],
    };

    if (!cell || !dayColumns.includes(col) || !validatedRows.includes(row)) {
      return summary;
    }

    const currentValue = sanitizeDisplayText(cell.textContent || "", {
      trim: true,
    });

    clearValidationState(cell);

    if (!currentValue) {
      rememberCellValidValue(cell, "");
      return summary;
    }

    if (!parsedMemberList.length) {
      const fallbackValue = sanitizeDisplayText(cell.dataset.lastValidValue || "", {
        trim: true,
      });

      cell.textContent = fallbackValue;
      flashCellState(cell, "cell-invalid");
      summary.isValid = false;

      if (announce) {
        updateStatusMessage(
          "Hãy dán hoặc tải danh sách trước, rồi mới nhập tên vào các ô công việc."
        );
        summary.announced = true;
      }

      return summary;
    }

    const canonicalName = getCanonicalMemberName(currentValue);

    if (!canonicalName) {
      const fallbackValue = sanitizeDisplayText(cell.dataset.lastValidValue || "", {
        trim: true,
      });

      cell.textContent = fallbackValue;
      flashCellState(cell, "cell-invalid");
      summary.isValid = false;

      if (announce) {
        updateStatusMessage(
          `"${currentValue}" không có trong danh sách đã tải lên nên ô này đã được ${fallbackValue ? "trả về tên hợp lệ trước đó" : "xóa"}.`
        );
        summary.announced = true;
      }

      return summary;
    }

    if (currentValue !== canonicalName) {
      cell.textContent = canonicalName;
    }

    rememberCellValidValue(cell, canonicalName);

    if (sourceRows.includes(row)) {
      const duplicateSummary = autoReplaceDuplicateSourceMembers(col, row);
      summary.replacedRows = duplicateSummary.replacedRows;
      summary.clearedRows = duplicateSummary.clearedRows;

      if (announce && (summary.replacedRows.length || summary.clearedRows.length)) {
        const replacedNames = summary.replacedRows
          .map((item) => item.replacementName)
          .join(", ");

        if (summary.replacedRows.length && summary.clearedRows.length) {
          updateStatusMessage(
            `${getColumnTitle(col)} bị trùng tên ở hàng 2–8 nên hệ thống đã tự đổi các ô còn lại sang ${replacedNames}; những ô chưa còn người phù hợp thì để trống.`
          );
        } else if (summary.replacedRows.length) {
          updateStatusMessage(
            `${getColumnTitle(col)} bị trùng tên ở hàng 2–8 nên hệ thống đã tự đổi các ô trùng còn lại sang ${replacedNames}.`
          );
        } else {
          updateStatusMessage(
            `${getColumnTitle(col)} bị trùng tên ở hàng 2–8 nhưng danh sách ngoài không còn ai chưa dùng, nên các ô trùng còn lại đã được để trống.`
          );
        }

        summary.announced = true;
      }
    }

    return summary;
  }

  function reconcileManualEntriesWithMemberList() {
    dayColumns.forEach((col) => {
      validatedRows.forEach((row) => {
        const cell = getCell(row, col);
        const displayValue = getCellDisplayValue(row, col);

        if (!cell) return;

        clearValidationState(cell);

        if (!displayValue) {
          rememberCellValidValue(cell, "");
          return;
        }

        if (!parsedMemberList.length) {
          return;
        }

        const canonicalName = getCanonicalMemberName(displayValue);

        if (!canonicalName) {
          setCellValue(row, col, "");
          rememberCellValidValue(cell, "");
          flashCellState(cell, "cell-invalid");
          return;
        }

        setCellValue(row, col, canonicalName, { rememberValid: true });
      });

      autoReplaceDuplicateSourceMembers(col, null);
    });
  }

  function parseMemberList(rawText = "") {
    const pieces = rawText.split(/[\n,;\t]+/g);
    const uniqueNames = [];
    const seen = new Set();

    pieces.forEach((piece) => {
      const displayName = sanitizeDisplayText(piece, { trim: true });
      const memberKey = normalizeMemberKey(displayName);

      if (!memberKey || seen.has(memberKey)) return;
      seen.add(memberKey);
      uniqueNames.push(displayName);
    });

    return uniqueNames;
  }

  function shuffleArray(array) {
    const next = [...array];
    for (let i = next.length - 1; i > 0; i -= 1) {
      const randomIndex = Math.floor(Math.random() * (i + 1));
      [next[i], next[randomIndex]] = [next[randomIndex], next[i]];
    }
    return next;
  }

  function buildRandomSequence(pool, targetLength) {
    if (!Array.isArray(pool) || pool.length === 0) {
      return Array.from({ length: targetLength }, () => "");
    }

    const result = [];
    while (result.length < targetLength) {
      result.push(...shuffleArray(pool));
    }

    return result.slice(0, targetLength);
  }

  function updateListSummary() {
    if (memberListCount) {
      memberListCount.textContent = `${parsedMemberList.length} tên hợp lệ`;
    }
  }

  function updateStatusMessage(text) {
    if (memberListStatus) {
      memberListStatus.textContent = text;
    }
  }

  function getColumnTitle(col) {
    return getCellDisplayValue(1, col) || `Cột ${col - 3}`;
  }

  function setActiveColumn(col) {
    if (!dayColumns.includes(col)) return;
    activeColumn = col;
  }

  function getEditableDayCellsInOrder() {
    const orderedCells = [];

    dayColumns.forEach((col) => {
      [1, ...manualRows].forEach((row) => {
        const cell = getCell(row, col);
        if (cell && cell.contentEditable === "true") {
          orderedCells.push(cell);
        }
      });
    });

    return orderedCells;
  }

  function focusEditableCell(cell) {
    if (!cell || cell.contentEditable !== "true") return;

    cell.focus();

    const textNode = cell.firstChild;
    if (!textNode) return;

    const range = document.createRange();
    const selection = window.getSelection();
    range.selectNodeContents(cell);
    range.collapse(false);
    selection.removeAllRanges();
    selection.addRange(range);
  }

  function moveFocusToNextEditableCell(currentCell) {
    const orderedCells = getEditableDayCellsInOrder();
    const currentIndex = orderedCells.indexOf(currentCell);

    if (currentIndex === -1) return false;

    const nextCell = orderedCells[currentIndex + 1];
    if (!nextCell) return false;

    focusEditableCell(nextCell);
    return true;
  }

  function clearGeneratedRowsForColumn(col) {
    [...permuteRows, ...randomPoolRows].forEach((row) => {
      setCellValue(row, col, "");
    });
  }

  function getSourceMembersForColumn(col) {
    return sourceRows.map((row) => getCellDisplayValue(row, col));
  }

  function isColumnComplete(col) {
    return getSourceMembersForColumn(col).every(Boolean);
  }

  function applyAssignmentsForColumn(col) {
    const sourceMembers = getSourceMembersForColumn(col);

    if (!parsedMemberList.length || !sourceMembers.every(Boolean)) {
      clearGeneratedRowsForColumn(col);
      return {
        isReady: false,
        sourceCount: sourceMembers.filter(Boolean).length,
        eligibleCount: 0,
      };
    }

    const sourceSet = new Set(sourceMembers.map((name) => normalizeMemberKey(name)));
    const shuffledMembers = shuffleArray(sourceMembers);
    const eligibleMembers = parsedMemberList.filter(
      (name) => !sourceSet.has(normalizeMemberKey(name))
    );
    const randomAssignments = buildRandomSequence(
      eligibleMembers,
      randomPoolRows.length
    );

    permuteRows.forEach((row, index) => {
      setCellValue(row, col, shuffledMembers[index] || "");
    });

    randomPoolRows.forEach((row, index) => {
      setCellValue(row, col, randomAssignments[index] || "");
    });

    return {
      isReady: true,
      sourceCount: sourceMembers.length,
      eligibleCount: eligibleMembers.length,
    };
  }

  function refreshColumn(col, { announce = true } = {}) {
    if (!dayColumns.includes(col)) return null;

    const columnTitle = getColumnTitle(col);

    if (!parsedMemberList.length) {
      clearGeneratedRowsForColumn(col);
      if (announce) {
        updateStatusMessage(
          `Chưa có danh sách ngoài, nên ${columnTitle} chưa hiển thị các nhóm bên dưới.`
        );
      }
      return {
        status: "missing-list",
        sourceCount: getSourceMembersForColumn(col).filter(Boolean).length,
      };
    }

    const sourceMembers = getSourceMembersForColumn(col);
    const sourceCount = sourceMembers.filter(Boolean).length;

    if (!sourceMembers.every(Boolean)) {
      clearGeneratedRowsForColumn(col);
      if (announce) {
        updateStatusMessage(
          `${columnTitle} mới có ${sourceCount}/7 tên ở hàng 2–8, nên các nhóm bên dưới của cột này vẫn để trống.`
        );
      }
      return {
        status: "incomplete",
        sourceCount,
      };
    }

    const summary = applyAssignmentsForColumn(col);

    if (announce && summary.eligibleCount === 0) {
      updateStatusMessage(
        `${columnTitle} đã xáo 7 tên xuống dòng 28–34, nhưng danh sách ngoài không còn ai sau khi loại trừ 7 tên của cột này nên các dòng 12–27 và 35–44 đang để trống.`
      );
    } else if (announce) {
      updateStatusMessage(
        `${columnTitle} đã hiển thị xếp nhóm bên dưới. Dòng 28–34 là 7 tên của hàng 2–8 được xáo ngẫu nhiên; dòng 12–27 và 35–44 lấy từ danh sách ngoài và không trùng 7 tên đó.`
      );
    }

    return {
      status: summary.eligibleCount === 0 ? "ready-empty-pool" : "ready",
      sourceCount: 7,
      eligibleCount: summary.eligibleCount,
    };
  }

  function refreshAllColumns({ announce = true } = {}) {
    let readyColumns = 0;
    let incompleteColumns = 0;

    dayColumns.forEach((col) => {
      const result = refreshColumn(col, { announce: false });
      if (!result) return;
      if (result.status === "ready" || result.status === "ready-empty-pool") {
        readyColumns += 1;
      } else if (result.status === "incomplete") {
        incompleteColumns += 1;
      }
    });

    if (!announce) return;

    if (!parsedMemberList.length) {
      updateStatusMessage(
        "Chưa có danh sách ngoài nên các nhóm bên dưới của từng cột vẫn đang để trống."
      );
      return;
    }

    if (readyColumns === 0) {
      updateStatusMessage(
        "Đã cập nhật danh sách ngoài, nhưng chưa có cột nào đủ 7 tên ở hàng 2–8 nên chưa hiển thị nhóm bên dưới."
      );
      return;
    }

    if (incompleteColumns > 0) {
      updateStatusMessage(
        `Đã cập nhật danh sách ngoài. Có ${readyColumns} cột đã đủ 7 tên nên đã hiển thị nhóm bên dưới; ${incompleteColumns} cột còn thiếu nên vẫn để trống.`
      );
      return;
    }

    updateStatusMessage(
      `Đã cập nhật danh sách ngoài. Cả ${readyColumns} cột đã đủ 7 tên ở hàng 2–8 nên đều đã hiển thị nhóm bên dưới theo từng cột riêng.`
    );
  }

  function scheduleRefreshColumn(col, { announce = true } = {}) {
    if (!dayColumns.includes(col)) return;

    const currentTimer = columnRefreshTimers.get(col);
    if (currentTimer) {
      window.clearTimeout(currentTimer);
    }

    columnRefreshTimers.set(
      col,
      window.setTimeout(() => {
        refreshColumn(col, { announce });
        columnRefreshTimers.delete(col);
      }, 120)
    );
  }

  function scheduleRefreshAllColumns({ announce = true } = {}) {
    if (listRefreshTimer) {
      window.clearTimeout(listRefreshTimer);
    }

    listRefreshTimer = window.setTimeout(() => {
      refreshAllColumns({ announce });
      listRefreshTimer = null;
    }, 120);
  }

  function refreshParsedList() {
    parsedMemberList = parseMemberList(memberListInput ? memberListInput.value : "");
    updateListSummary();
    reconcileManualEntriesWithMemberList();
    scheduleRefreshAllColumns({ announce: true });
  }

  function createScheduleTable(rows = 45, cols = 7) {
    if (!scheduleTableBody) return;

    cellRegistry.clear();
    scheduleTableBody.innerHTML = "";

    const highlightColor = "#B4C6E7";
    const thickBorder = "2px solid #000";

    const col1Width = "50px";
    const col2Width = "50px";
    const col3Width = "180px";

    const makeEditableCell = (
      cell,
      defaultText = "",
      { forceUppercaseDisplay = false, autoCapitalizeName = false } = {}
    ) => {
      cell.contentEditable = "true";
      cell.spellcheck = false;
      cell.style.cursor = "text";
      cell.style.textAlign = "center";
      cell.style.verticalAlign = "middle";
      cell.style.textTransform = forceUppercaseDisplay ? "uppercase" : "none";
      cell.style.whiteSpace = "pre-wrap";
      cell.textContent = sanitizeDisplayText(defaultText, { trim: true });

      cell.addEventListener("keydown", function (event) {
        if (event.key === "Enter") {
          event.preventDefault();

          const row = Number(cell.dataset.row || 0);
          const col = Number(cell.dataset.col || 0);
          let shouldMoveNext = true;

          if (dayColumns.includes(col) && validatedRows.includes(row)) {
            const validationSummary = validateManualEntryCell(cell, {
              announce: true,
            });

            if (sourceRows.includes(row)) {
              scheduleRefreshColumn(col, {
                announce: !validationSummary.announced,
              });
            }

            shouldMoveNext = validationSummary.isValid;
          }

          if (shouldMoveNext) {
            moveFocusToNextEditableCell(cell);
          }
        }
      });

      cell.addEventListener("input", function () {
        const selection = window.getSelection();
        let caretOffset = 0;

        if (selection && selection.rangeCount > 0) {
          const range = selection.getRangeAt(0);
          const preCaretRange = range.cloneRange();
          preCaretRange.selectNodeContents(cell);
          preCaretRange.setEnd(range.endContainer, range.endOffset);
          caretOffset = preCaretRange.toString().length;
        }

        let normalizedText = sanitizeDisplayText(cell.textContent, {
          trim: false,
        });

        if (autoCapitalizeName) {
          normalizedText = formatVietnameseNameCase(normalizedText);
        }

        if (cell.textContent !== normalizedText) {
          cell.textContent = normalizedText;

          if (cell.firstChild) {
            const newRange = document.createRange();
            const newSelection = window.getSelection();
            const safeOffset = Math.min(
              caretOffset,
              cell.firstChild.textContent.length
            );

            newRange.setStart(cell.firstChild, safeOffset);
            newRange.collapse(true);
            newSelection.removeAllRanges();
            newSelection.addRange(newRange);
          }
        }
      });

      cell.addEventListener("blur", function () {
        let finalText = sanitizeDisplayText(cell.textContent, { trim: true });

        if (autoCapitalizeName) {
          finalText = formatVietnameseNameCase(finalText);
        }

        if (cell.textContent !== finalText) {
          cell.textContent = finalText;
        }
      });

      cell.addEventListener("paste", function (event) {
        event.preventDefault();

        let pastedText = sanitizeDisplayText(
          (event.clipboardData || window.clipboardData).getData("text"),
          { trim: true }
        );

        if (autoCapitalizeName) {
          pastedText = formatVietnameseNameCase(pastedText);
        }

        const selection = window.getSelection();

        if (!selection || selection.rangeCount === 0) {
          cell.textContent += pastedText;
          return;
        }

        const range = selection.getRangeAt(0);
        range.deleteContents();
        range.insertNode(document.createTextNode(pastedText));
        range.collapse(false);

        selection.removeAllRanges();
        selection.addRange(range);

        cell.dispatchEvent(new Event("input", { bubbles: true }));
      });
    };

    const applyFullTableBorder = (cell, rowIndex, colIndex) => {
      if (rowIndex === 1) {
        cell.style.borderTop = thickBorder;
      }

      if (rowIndex === rows) {
        cell.style.borderBottom = thickBorder;
      }

      if (colIndex === 1) {
        cell.style.borderLeft = thickBorder;
      }

      if (colIndex === cols) {
        cell.style.borderRight = thickBorder;
      }

      if ([11, 19, 27, 34].includes(rowIndex)) {
        cell.style.borderBottom = thickBorder;
      }

      if ([12, 20, 28, 35].includes(rowIndex)) {
        cell.style.borderTop = thickBorder;
      }
    };

    const createStaticLabelCell = ({ text = "", html = "", rowSpan = 1 }) => {
      const cell = document.createElement("td");
      cell.colSpan = 3;
      cell.rowSpan = rowSpan;
      cell.style.textAlign = "center";
      cell.style.verticalAlign = "middle";
      cell.style.padding = "0 6px";
      cell.style.textTransform = "uppercase";
      cell.style.whiteSpace = "pre-wrap";

      if (html) {
        cell.innerHTML = html;
      } else {
        cell.textContent = text;
      }

      return cell;
    };

    const createVerticalWordsCell = ({
      text = "",
      rowSpan = 1,
      colSpan = 1,
      fontWeight = "normal",
    }) => {
      const cell = document.createElement("td");
      cell.rowSpan = rowSpan;
      cell.colSpan = colSpan;
      cell.style.textAlign = "center";
      cell.style.verticalAlign = "middle";
      cell.style.padding = "4px 2px";
      cell.style.textTransform = "uppercase";
      cell.style.whiteSpace = "pre-wrap";
      cell.style.fontWeight = fontWeight;
      cell.style.lineHeight = "1.4";

      const words = text.split(" ");
      cell.innerHTML = words.join("<br>");

      return cell;
    };

    const createMultilineInfoCell = ({ html = "", rowSpan = 1 }) => {
      const cell = document.createElement("td");
      cell.rowSpan = rowSpan;
      cell.style.textAlign = "center";
      cell.style.verticalAlign = "middle";
      cell.style.padding = "6px";
      cell.style.whiteSpace = "pre-wrap";
      cell.style.lineHeight = "1.5";
      cell.innerHTML = html;
      return cell;
    };

    for (let row = 1; row <= rows; row += 1) {
      const tr = document.createElement("tr");

      const applyCellStyle = (cell, colIndex) => {
        cell.style.textAlign = "center";
        cell.style.verticalAlign = "middle";

        const isDayColumn = dayColumns.includes(colIndex);
        const isManualAssignmentCell = isDayColumn && manualRows.includes(row);
        const isGeneratedAssignmentCell = isDayColumn && generatedRows.includes(row);

        if (isDayColumn) {
          registerCell(row, colIndex, cell);
        }

        if (row === 1) {
          cell.style.backgroundColor = highlightColor;
          cell.style.fontWeight = "bold";

          if (colIndex === 1) {
            cell.textContent = "CÔNG TÁC";
            cell.contentEditable = "false";
          }

          if (isDayColumn) {
            makeEditableCell(cell, "Chúa nhật", {
              forceUppercaseDisplay: true,
            });
          }
        }

        if ([12, 20, 28, 35].includes(row) && isDayColumn) {
          cell.style.backgroundColor = highlightColor;
          cell.style.fontWeight = "bold";
        }

        if (isManualAssignmentCell && cell.contentEditable !== "true") {
          makeEditableCell(cell, "", {
            forceUppercaseDisplay: false,
            autoCapitalizeName: true,
          });
          cell.classList.add("manual-entry-cell");
        } else if (isGeneratedAssignmentCell && cell.contentEditable !== "true") {
          makeEditableCell(cell, "", {
            forceUppercaseDisplay: false,
            autoCapitalizeName: true,
          });
          cell.classList.add("manual-entry-cell", "generated-entry-cell");
        } else if (isDayColumn && cell.contentEditable !== "true") {
          cell.contentEditable = "false";
          cell.spellcheck = false;
        }

        if (randomPoolRows.includes(row) && isDayColumn) {
          cell.classList.add("auto-fill-cell", "generated-entry-cell");
        }

        if (permuteRows.includes(row) && isDayColumn) {
          cell.classList.add("auto-permute-cell", "generated-entry-cell");
        }

        applyFullTableBorder(cell, row, colIndex);
      };

      if (row === 1) {
        const mergedCell = document.createElement("td");
        mergedCell.colSpan = 3;
        applyCellStyle(mergedCell, 1);
        tr.appendChild(mergedCell);

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row === 2) {
        const cell = createStaticLabelCell({
          html: "Xướng kinh - <strong>Chính</strong>",
        });
        applyFullTableBorder(cell, row, 1);
        tr.appendChild(cell);

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row === 3) {
        const cell = createStaticLabelCell({
          html: "Xướng kinh - <strong>Phụ</strong>",
        });
        applyFullTableBorder(cell, row, 1);
        tr.appendChild(cell);

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row === 4) {
        const cell = createStaticLabelCell({
          text: "Bài đọc Kinh Sách 1",
        });
        applyFullTableBorder(cell, row, 1);
        tr.appendChild(cell);

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row === 5) {
        const cell = createStaticLabelCell({
          text: "Bài đọc Kinh Sách 2",
        });
        applyFullTableBorder(cell, row, 1);
        tr.appendChild(cell);

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row === 6) {
        const cell = createStaticLabelCell({
          text: "Trực chuông, Giúp lễ",
          rowSpan: 2,
        });
        applyFullTableBorder(cell, row, 1);
        tr.appendChild(cell);

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row === 7) {
        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row === 8) {
        const cell = createStaticLabelCell({
          text: "Ca trưởng",
        });
        applyFullTableBorder(cell, row, 1);
        tr.appendChild(cell);

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row === 9) {
        const cell = createStaticLabelCell({
          text: "Đàn",
        });
        applyFullTableBorder(cell, row, 1);
        tr.appendChild(cell);

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row === 10) {
        const cell = createStaticLabelCell({
          text: "Đọc sách giờ cơm",
          rowSpan: 2,
        });
        applyFullTableBorder(cell, row, 1);
        tr.appendChild(cell);

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row === 11) {
        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row >= 12 && row <= 19) {
        if (row === 12) {
          const cell1 = createVerticalWordsCell({
            text: "Dọn cơm & Vệ Sinh Thỉnh viện",
            rowSpan: 16,
          });
          applyFullTableBorder(cell1, row, 1);
          cell1.style.borderBottom = thickBorder;
          cell1.style.width = col1Width;
          cell1.style.maxWidth = col1Width;
          tr.appendChild(cell1);

          const cell2 = createVerticalWordsCell({
            text: "Nhóm 1",
            rowSpan: 8,
          });
          cell2.style.borderTop = thickBorder;
          cell2.style.width = col2Width;
          cell2.style.maxWidth = col2Width;
          tr.appendChild(cell2);

          const cell3 = createMultilineInfoCell({
            rowSpan: 8,
            html:
              "<strong>Dọn cơm:</strong><br>Tối T7 → Sáng T4<br><br><strong>Vệ sinh sáng Thỉnh viện:</strong><br>Sáng T5 → Sáng T7",
          });
          cell3.style.borderTop = thickBorder;
          cell3.style.width = col3Width;
          tr.appendChild(cell3);
        }

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row >= 20 && row <= 27) {
        if (row === 20) {
          const cell2 = createVerticalWordsCell({
            text: "Nhóm 2",
            rowSpan: 8,
          });
          cell2.style.borderTop = thickBorder;
          cell2.style.width = col2Width;
          cell2.style.maxWidth = col2Width;
          tr.appendChild(cell2);

          const cell3 = createMultilineInfoCell({
            rowSpan: 8,
            html:
              "<strong>Dọn cơm:</strong><br>Trưa T4 → Trưa T7<br><br><strong>Vệ sinh sáng Thỉnh viện:</strong><br>Sáng CN → Sáng T4",
          });
          cell3.style.borderTop = thickBorder;
          cell3.style.borderBottom = thickBorder;
          cell3.style.width = col3Width;
          tr.appendChild(cell3);
        }

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row >= 28 && row <= 34) {
        if (row === 28) {
          const cell = createStaticLabelCell({
            html:
              '<div style="display:flex;align-items:center;justify-content:center;height:100%;text-align:center;line-height:1.6;"><div><strong>Sáng:</strong> Tưới cây<br>Vệ sinh tu viện<br><strong>Chiều:</strong> Tưới cây<br>Dọn Nhà Nguyện</div></div>',
            rowSpan: 7,
          });
          applyFullTableBorder(cell, row, 1);
          tr.appendChild(cell);
        }

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row >= 35 && row <= 44) {
        if (row === 35) {
          const cell = createStaticLabelCell({
            html:
              '<div style="display:flex;align-items:center;justify-content:center;height:100%;text-align:center;line-height:1.6;"><div><strong>Sáng:</strong> Quét sân<br><strong>Trưa:</strong> Dọn WC trệt<br><strong>Chiều:</strong> Quét sân</div></div>',
            rowSpan: 10,
          });
          applyFullTableBorder(cell, row, 1);
          tr.appendChild(cell);
        }

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else if (row === 45) {
        const mergedCell = document.createElement("td");
        mergedCell.colSpan = 3;
        applyFullTableBorder(mergedCell, row, 1);
        tr.appendChild(mergedCell);

        for (let col = 4; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      } else {
        for (let col = 1; col <= cols; col += 1) {
          const td = document.createElement("td");
          applyCellStyle(td, col);
          tr.appendChild(td);
        }
      }

      scheduleTableBody.appendChild(tr);
    }
  }

  function placeCaretAtEnd(element) {
    if (!element) return;

    element.focus();

    const range = document.createRange();
    const selection = window.getSelection();

    range.selectNodeContents(element);
    range.collapse(false);

    selection.removeAllRanges();
    selection.addRange(range);
  }

  function getCaretOffsetWithin(element) {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return 0;

    const range = selection.getRangeAt(0);
    const preCaretRange = range.cloneRange();
    preCaretRange.selectNodeContents(element);
    preCaretRange.setEnd(range.endContainer, range.endOffset);

    return preCaretRange.toString().length;
  }

  function setCaretOffsetWithin(element, offset) {
    const selection = window.getSelection();
    const range = document.createRange();

    if (!element.firstChild) {
      element.textContent = "";
    }

    const textNode = element.firstChild;
    const safeOffset = Math.min(offset, textNode.textContent.length);

    range.setStart(textNode, safeOffset);
    range.collapse(true);

    selection.removeAllRanges();
    selection.addRange(range);
  }

  function insertTextAtCursor(text) {
    const selection = window.getSelection();

    if (!selection || selection.rangeCount === 0) {
      editableSuffix.textContent += text;
      placeCaretAtEnd(editableSuffix);
      return;
    }

    const range = selection.getRangeAt(0);
    range.deleteContents();
    range.insertNode(document.createTextNode(text));
    range.collapse(false);

    selection.removeAllRanges();
    selection.addRange(range);
  }

  function normalizeEditableSuffix({ trim = false } = {}) {
    if (!editableSuffix) return;

    const isFocused = document.activeElement === editableSuffix;
    const caretOffset = isFocused ? getCaretOffsetWithin(editableSuffix) : 0;

    const normalizedText = sanitizeDisplayText(editableSuffix.textContent, {
      trim,
    });

    if (editableSuffix.textContent !== normalizedText) {
      editableSuffix.textContent = normalizedText;

      if (isFocused) {
        setCaretOffsetWithin(editableSuffix, caretOffset);
      }
    }
  }

  if (titleLine && editableSuffix) {
    titleLine.addEventListener("dblclick", function () {
      placeCaretAtEnd(editableSuffix);
    });

    editableSuffix.addEventListener("input", function () {
      normalizeEditableSuffix({ trim: false });
    });

    editableSuffix.addEventListener("keydown", function (event) {
      if (event.key === "Enter") {
        event.preventDefault();
      }
    });

    editableSuffix.addEventListener("blur", function () {
      normalizeEditableSuffix({ trim: true });
    });

    editableSuffix.addEventListener("paste", function (event) {
      event.preventDefault();

      const pastedText = sanitizeDisplayText(
        (event.clipboardData || window.clipboardData).getData("text"),
        { trim: true }
      );

      insertTextAtCursor(pastedText);
      normalizeEditableSuffix({ trim: false });
    });
  }

  function setButtonLoading(button, isLoading, textLoading, textNormal) {
    if (!button) return;

    if (isLoading) {
      button.disabled = true;
      button.dataset.originalText = button.textContent;
      button.textContent = textLoading;
    } else {
      button.disabled = false;
      button.textContent =
        textNormal || button.dataset.originalText || button.textContent;
    }
  }

  async function waitForImages(container) {
    const images = Array.from(container.querySelectorAll("img"));

    await Promise.all(
      images.map((img) => {
        if (img.complete && img.naturalWidth > 0) {
          return Promise.resolve();
        }

        return new Promise((resolve, reject) => {
          const onLoad = () => {
            cleanup();
            resolve();
          };

          const onError = () => {
            cleanup();
            reject(new Error("Không tải được ảnh trong trang."));
          };

          const cleanup = () => {
            img.removeEventListener("load", onLoad);
            img.removeEventListener("error", onError);
          };

          img.addEventListener("load", onLoad);
          img.addEventListener("error", onError);
        });
      })
    );
  }

  async function createCanvasFromElement(element, options = {}) {
    if (typeof html2canvas !== "function") {
      throw new Error("Thiếu thư viện html2canvas.");
    }

    if (document.activeElement) {
      document.activeElement.blur();
    }

    await waitForImages(element);

    const elementWidth = Math.ceil(options.width || element.scrollWidth || element.offsetWidth || 0);
    const elementHeight = Math.ceil(options.height || element.scrollHeight || element.offsetHeight || 0);

    return await html2canvas(element, {
      scale: options.scale || 2,
      useCORS: true,
      allowTaint: true,
      backgroundColor: "#ffffff",
      scrollX: 0,
      scrollY: 0,
      width: elementWidth,
      height: elementHeight,
      windowWidth: options.windowWidth || elementWidth,
      windowHeight: options.windowHeight || elementHeight,
    });
  }

  function applyPdfLayoutPreset(clone, preset) {
    Object.assign(clone.style, {
      width: "100%",
      maxWidth: "100%",
      margin: "0",
      padding: "0",
      borderRadius: "0",
      boxShadow: "none",
      background: "#ffffff",
      boxSizing: "border-box",
      position: "absolute",
      left: "0",
      top: "0",
      transformOrigin: "top left",
    });

    const header = clone.querySelector(".header");
    if (header) {
      header.style.margin = `0 0 ${preset.headerGap}px`;
      header.style.alignItems = "flex-start";
    }

    const logo = clone.querySelector(".logo");
    if (logo) {
      logo.style.width = `${preset.logoWidth}px`;
      logo.style.marginRight = `${preset.logoGap}px`;
      logo.style.flexShrink = "0";
    }

    clone.querySelectorAll(".institution-info h3").forEach((node) => {
      node.style.fontSize = `${preset.headerTitleFont}px`;
      node.style.lineHeight = String(preset.headerTitleLineHeight);
    });

    clone.querySelectorAll(".institution-info p").forEach((node) => {
      node.style.margin = `${preset.headerBodyMarginY}px 0`;
      node.style.fontSize = `${preset.headerBodyFont}px`;
      node.style.lineHeight = String(preset.headerBodyLineHeight);
    });

    const title = clone.querySelector(".title-line");
    if (title) {
      title.style.marginBottom = `${preset.titleGap}px`;
      title.style.fontSize = `${preset.titleFont}px`;
      title.style.lineHeight = String(preset.titleLineHeight);
      title.style.whiteSpace = "normal";
    }

    const tableWrap = clone.querySelector(".schedule-table-wrap");
    if (tableWrap) {
      tableWrap.style.marginTop = "0";
    }

    const table = clone.querySelector(".schedule-table");
    if (table) {
      table.style.width = "100%";
      table.style.tableLayout = "fixed";
      table.style.borderCollapse = "collapse";
      table.style.fontSize = `${preset.tableFont}px`;
    }

    clone.querySelectorAll(".schedule-table td").forEach((cell) => {
      const currentWidth = (cell.style.width || "").trim();

      cell.style.height = `${preset.rowHeight}px`;
      cell.style.padding = `${preset.cellPadY}px ${preset.cellPadX}px`;
      cell.style.lineHeight = String(preset.cellLineHeight);
      cell.style.fontSize = `${preset.cellFont}px`;
      cell.style.wordBreak = "break-word";
      cell.style.overflowWrap = "anywhere";

      if (currentWidth === "50px") {
        cell.style.width = `${preset.narrowColWidth}px`;
        cell.style.maxWidth = `${preset.narrowColWidth}px`;
      } else if (currentWidth === "180px") {
        cell.style.width = `${preset.infoColWidth}px`;
        cell.style.maxWidth = `${preset.infoColWidth}px`;
      }
    });

    clone.querySelectorAll('.schedule-table td[colspan="3"]').forEach((cell) => {
      cell.style.fontSize = `${preset.sideFont}px`;
      cell.style.lineHeight = String(preset.sideLineHeight);
      cell.style.padding = `0 ${preset.sidePadX}px`;
    });

    clone.querySelectorAll(".schedule-table td[rowspan]").forEach((cell) => {
      cell.style.fontSize = `${preset.sideFont}px`;
      cell.style.lineHeight = String(preset.sideLineHeight);
      cell.style.padding = `${preset.sidePadY}px ${preset.sidePadX}px`;
    });

    clone.querySelectorAll('.schedule-table td[data-row="1"]').forEach((cell) => {
      cell.style.fontSize = `${preset.dayHeaderFont}px`;
      cell.style.lineHeight = String(preset.dayHeaderLineHeight);
    });

    clone.querySelectorAll('.schedule-table td[data-col="4"], .schedule-table td[data-col="5"], .schedule-table td[data-col="6"], .schedule-table td[data-col="7"]').forEach((cell) => {
      cell.style.fontSize = `${preset.nameFont}px`;
      cell.style.lineHeight = String(preset.nameLineHeight);
      cell.style.padding = `0 ${preset.namePadX}px`;
    });
  }

  function stretchPdfTableHeight(clone, extraNaturalHeight) {
    if (!clone || !(extraNaturalHeight > 0)) return;

    const rows = Array.from(clone.querySelectorAll(".schedule-table tbody tr"));
    if (!rows.length) return;

    const extraPerRow = extraNaturalHeight / rows.length;

    rows.forEach((row, index) => {
      const measuredHeight = row.getBoundingClientRect().height || row.offsetHeight || 0;
      if (!measuredHeight) return;

      const adjustedHeight = measuredHeight + extraPerRow + (index === rows.length - 1 ? 0.5 : 0);
      row.style.height = `${adjustedHeight}px`;
      row.style.minHeight = `${adjustedHeight}px`;
    });
  }

  function createPdfCaptureClone() {
    const wrapper = document.createElement("div");
    wrapper.className = "pdf-export-host";

    Object.assign(wrapper.style, {
      position: "fixed",
      left: "0",
      top: "0",
      width: "210mm",
      height: "297mm",
      margin: "0",
      padding: "0",
      background: "#ffffff",
      pointerEvents: "none",
      zIndex: "-1",
      overflow: "hidden",
      opacity: "0",
    });

    const page = document.createElement("div");
    page.className = "pdf-export-page";

    Object.assign(page.style, {
      width: "210mm",
      height: "297mm",
      margin: "0",
      padding: "0",
      boxSizing: "border-box",
      background: "#ffffff",
      overflow: "hidden",
      position: "relative",
    });

    const content = document.createElement("div");
    content.className = "pdf-export-content";

    Object.assign(content.style, {
      width: "208mm",
      height: "295mm",
      margin: "1mm auto",
      position: "relative",
      overflow: "hidden",
      background: "#ffffff",
      boxSizing: "border-box",
    });

    const clone = captureArea.cloneNode(true);
    clone.removeAttribute("id");
    clone.classList.add("pdf-export-mode");

    content.appendChild(clone);
    page.appendChild(content);
    wrapper.appendChild(page);
    document.body.appendChild(wrapper);

    const presets = [
      {
        headerGap: 8,
        logoWidth: 54,
        logoGap: 10,
        headerTitleFont: 9,
        headerTitleLineHeight: 1.18,
        headerBodyFont: 7.2,
        headerBodyLineHeight: 1.12,
        headerBodyMarginY: 1,
        titleGap: 8,
        titleFont: 16,
        titleLineHeight: 1.15,
        tableFont: 8.4,
        cellFont: 7.7,
        dayHeaderFont: 8.9,
        dayHeaderLineHeight: 1.05,
        nameFont: 7.9,
        nameLineHeight: 1.03,
        namePadX: 2,
        rowHeight: 17.6,
        cellPadY: 0,
        cellPadX: 1,
        cellLineHeight: 1.03,
        sideFont: 7.9,
        sideLineHeight: 1.08,
        sidePadX: 2,
        sidePadY: 1,
        narrowColWidth: 34,
        infoColWidth: 118,
      },
      {
        headerGap: 6,
        logoWidth: 48,
        logoGap: 8,
        headerTitleFont: 8.4,
        headerTitleLineHeight: 1.14,
        headerBodyFont: 6.8,
        headerBodyLineHeight: 1.08,
        headerBodyMarginY: 1,
        titleGap: 6,
        titleFont: 15,
        titleLineHeight: 1.1,
        tableFont: 8,
        cellFont: 7.4,
        dayHeaderFont: 8.4,
        dayHeaderLineHeight: 1.02,
        nameFont: 7.6,
        nameLineHeight: 1.02,
        namePadX: 1,
        rowHeight: 16.8,
        cellPadY: 0,
        cellPadX: 1,
        cellLineHeight: 1.01,
        sideFont: 7.5,
        sideLineHeight: 1.05,
        sidePadX: 2,
        sidePadY: 1,
        narrowColWidth: 32,
        infoColWidth: 110,
      },
      {
        headerGap: 5,
        logoWidth: 44,
        logoGap: 7,
        headerTitleFont: 8,
        headerTitleLineHeight: 1.1,
        headerBodyFont: 6.5,
        headerBodyLineHeight: 1.06,
        headerBodyMarginY: 0,
        titleGap: 5,
        titleFont: 14,
        titleLineHeight: 1.05,
        tableFont: 7.6,
        cellFont: 7.1,
        dayHeaderFont: 8,
        dayHeaderLineHeight: 1,
        nameFont: 7.2,
        nameLineHeight: 1,
        namePadX: 1,
        rowHeight: 15.9,
        cellPadY: 0,
        cellPadX: 0,
        cellLineHeight: 1,
        sideFont: 7.1,
        sideLineHeight: 1.03,
        sidePadX: 1,
        sidePadY: 1,
        narrowColWidth: 30,
        infoColWidth: 102,
      },
    ];

    const availableWidth = Math.ceil(content.clientWidth || content.offsetWidth || 768);
    const availableHeight = Math.ceil(content.clientHeight || content.offsetHeight || 1096);
    let naturalWidth = availableWidth;
    let naturalHeight = availableHeight;

    let bestPreset = presets[presets.length - 1];
    for (const preset of presets) {
      applyPdfLayoutPreset(clone, preset);
      naturalWidth = Math.ceil(clone.scrollWidth || clone.offsetWidth || availableWidth);
      naturalHeight = Math.ceil(clone.scrollHeight || clone.offsetHeight || availableHeight);
      bestPreset = preset;
      if (naturalHeight <= availableHeight) {
        break;
      }
    }

    applyPdfLayoutPreset(clone, bestPreset);
    naturalWidth = Math.ceil(clone.scrollWidth || clone.offsetWidth || availableWidth);
    naturalHeight = Math.ceil(clone.scrollHeight || clone.offsetHeight || availableHeight);

    let scale = Math.min(availableWidth / naturalWidth, availableHeight / naturalHeight);
    let renderWidth = naturalWidth * scale;
    let renderHeight = naturalHeight * scale;
    const remainingHeight = availableHeight - renderHeight;

    if (remainingHeight > 16) {
      const extraNaturalHeight = Math.min(remainingHeight / Math.max(scale, 0.001), 110);
      stretchPdfTableHeight(clone, extraNaturalHeight);
      naturalWidth = Math.ceil(clone.scrollWidth || clone.offsetWidth || availableWidth);
      naturalHeight = Math.ceil(clone.scrollHeight || clone.offsetHeight || availableHeight);
      scale = Math.min(availableWidth / naturalWidth, availableHeight / naturalHeight);
      renderWidth = naturalWidth * scale;
      renderHeight = naturalHeight * scale;
    }

    const offsetX = Math.max((availableWidth - renderWidth) / 2, 0);
    const offsetY = Math.max(Math.min((availableHeight - renderHeight) / 3, 8), 0);

    clone.style.width = `${naturalWidth}px`;
    clone.style.left = `${offsetX}px`;
    clone.style.top = `${offsetY}px`;
    clone.style.transform = `scale(${scale})`;

    return {
      wrapper,
      page,
      content,
      clone,
      cleanup() {
        wrapper.remove();
      },
    };
  }

  function canvasToBlob(canvas, type = "image/png", quality = 1) {
    return new Promise((resolve, reject) => {
      canvas.toBlob(
        (blob) => {
          if (blob) {
            resolve(blob);
          } else {
            reject(new Error("Không tạo được dữ liệu ảnh."));
          }
        },
        type,
        quality
      );
    });
  }

  async function exportPdf() {
    if (!window.jspdf || !window.jspdf.jsPDF) {
      throw new Error("Thiếu thư viện jsPDF.");
    }

    const { jsPDF } = window.jspdf;
    const pdfCapture = createPdfCaptureClone();

    try {
      await new Promise((resolve) => requestAnimationFrame(resolve));
      await new Promise((resolve) => requestAnimationFrame(resolve));

      const canvas = await createCanvasFromElement(pdfCapture.page, {
        scale: 2.6,
        width: pdfCapture.page.offsetWidth,
        height: pdfCapture.page.offsetHeight,
        windowWidth: pdfCapture.page.offsetWidth,
        windowHeight: pdfCapture.page.offsetHeight,
      });

      const imgData = canvas.toDataURL("image/png", 1);

      const pdf = new jsPDF({
        orientation: "portrait",
        unit: "mm",
        format: "a4",
        compress: true,
      });

      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();

      pdf.addImage(imgData, "PNG", 0, 0, pageWidth, pageHeight, undefined, "FAST");
      pdf.save("lich-cong-tac-mua-a4.pdf");
    } finally {
      pdfCapture.cleanup();
    }
  }

  async function copyAsImage() {
    const canvas = await createCanvasFromElement(captureArea);
    const blob = await canvasToBlob(canvas);

    if (
      navigator.clipboard &&
      window.ClipboardItem &&
      typeof navigator.clipboard.write === "function"
    ) {
      await navigator.clipboard.write([
        new ClipboardItem({
          [blob.type]: blob,
        }),
      ]);
      return;
    }

    throw new Error("Trình duyệt không hỗ trợ sao chép ảnh trực tiếp.");
  }

  if (downloadPdfBtn) {
    downloadPdfBtn.addEventListener("click", async function () {
      try {
        setButtonLoading(downloadPdfBtn, true, "Đang tạo PDF...", "Tải PDF");
        await exportPdf();
      } catch (error) {
        alert(error.message || "Không thể xuất PDF.");
      } finally {
        setButtonLoading(downloadPdfBtn, false, "Đang tạo PDF...", "Tải PDF");
      }
    });
  }

  if (copyImageBtn) {
    copyImageBtn.addEventListener("click", async function () {
      try {
        setButtonLoading(
          copyImageBtn,
          true,
          "Đang sao chép...",
          "Sao chép ảnh"
        );
        await copyAsImage();
        alert("Đã sao chép ảnh vào clipboard.");
      } catch (error) {
        alert(
          error.message ||
            "Không thể sao chép ảnh. Hãy thử dùng trình duyệt khác."
        );
      } finally {
        setButtonLoading(
          copyImageBtn,
          false,
          "Đang sao chép...",
          "Sao chép ảnh"
        );
      }
    });
  }

  if (memberListInput) {
    memberListInput.addEventListener("input", refreshParsedList);
  }

  if (memberListFile) {
    memberListFile.addEventListener("change", async function (event) {
      const file = event.target.files && event.target.files[0];
      if (!file) return;

      try {
        const fileContent = await file.text();
        if (memberListInput) {
          memberListInput.value = fileContent;
        }
        refreshParsedList();
      } catch (error) {
        alert("Không đọc được file danh sách.");
      } finally {
        memberListFile.value = "";
      }
    });
  }

  if (randomizeAssignmentsBtn) {
    randomizeAssignmentsBtn.addEventListener("click", function () {
      if (!activeColumn) {
        updateStatusMessage("Chọn một cột rồi bấm ‘Random cột chọn’.");
        return;
      }

      refreshColumn(activeColumn, { announce: true });
    });
  }

  if (clearListBtn) {
    clearListBtn.addEventListener("click", function () {
      if (memberListInput) {
        memberListInput.value = "";
      }
      parsedMemberList = [];
      updateListSummary();
      refreshAllColumns({ announce: true });
    });
  }

  if (scheduleTableBody) {
    scheduleTableBody.addEventListener("focusin", function (event) {
      const rawTarget = event.target;
      const cell = rawTarget && rawTarget.closest
        ? rawTarget.closest("td[data-row][data-col]")
        : null;

      if (!cell) return;

      const row = Number(cell.dataset.row || 0);
      const col = Number(cell.dataset.col || 0);

      if (!dayColumns.includes(col)) return;

      setActiveColumn(col);

      if (validatedRows.includes(row)) {
        const displayValue = getCellDisplayValue(row, col);
        const canonicalName = getCanonicalMemberName(displayValue) || displayValue;
        rememberCellValidValue(cell, canonicalName);
      }
    });

    scheduleTableBody.addEventListener("click", function (event) {
      const rawTarget = event.target;
      const cell = rawTarget && rawTarget.closest
        ? rawTarget.closest("td[data-row][data-col]")
        : null;

      if (!cell) return;

      const col = Number(cell.dataset.col || 0);
      if (dayColumns.includes(col)) {
        setActiveColumn(col);
      }
    });

    scheduleTableBody.addEventListener("input", function (event) {
      const rawTarget = event.target;
      const cell = rawTarget && rawTarget.closest
        ? rawTarget.closest("td[data-row][data-col]")
        : null;

      if (!cell) return;

      const row = Number(cell.dataset.row || 0);
      const col = Number(cell.dataset.col || 0);

      if (!dayColumns.includes(col)) return;
      setActiveColumn(col);

      if (validatedRows.includes(row)) {
        clearValidationState(cell);
      }
    });

    scheduleTableBody.addEventListener("focusout", function (event) {
      const rawTarget = event.target;
      const cell = rawTarget && rawTarget.closest
        ? rawTarget.closest("td[data-row][data-col]")
        : null;

      if (!cell) return;

      const row = Number(cell.dataset.row || 0);
      const col = Number(cell.dataset.col || 0);

      if (!dayColumns.includes(col) || !validatedRows.includes(row)) return;

      const validationSummary = validateManualEntryCell(cell, { announce: true });

      if (sourceRows.includes(row)) {
        scheduleRefreshColumn(col, { announce: !validationSummary.announced });
      }
    });
  }

  createScheduleTable();
  updateListSummary();
  setActiveColumn(activeColumn);
  updateStatusMessage(
    "Dán hoặc tải danh sách ngoài trước. Tất cả ô tên ở cột 4–7 đều có thể sửa tay; tên nhập phải nằm trong danh sách này. Riêng hàng 2–8, nếu bị trùng tên thì hệ thống sẽ tự đổi ô trùng còn lại sang người chưa có."
  );
});
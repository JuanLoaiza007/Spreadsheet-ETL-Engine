/**
 * @fileoverview Spreadsheet ETL Engine
 * @version 1.0.4-rc1
 * @author JuanLoaiza007
 * @license MIT
 * @description A dynamic engine for extracting, transforming, and loading data within
 * Google Sheets. It features a custom expressive language for rule-based processing,
 * built-in syntax validation, and automated error handling.
 */

// --- Global configuration ---
const PREFIX = {
  FILTER: "_filter:",
  COMMENT_LINE: "//",
  EVAL: "eval:",
  CONSTANT: "constant:",
  FORMULA: "formula:",
  SRC: "src",
  SELF: "self",
};

// Symbols and operators defined in explicit form
const SYMBOLS = {
  OPEN: "[",
  CLOSE: "]",
  OR: "||",
  OP_EQUAL: "==",
  OP_NOT_EQUAL: "!=",
  OP_GREATER_EQUAL: ">=",
  OP_LESS_EQUAL: "<=",
  OP_GREATER: ">",
  OP_LESS: "<",
};

/**
 * Main function that executes the transformation process.
 */
function runMapping(config) {
  const ui = SpreadsheetApp.getUi();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Iniciando mapeo, espera un momento.`,
    "ETL Engine",
    4,
  );

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(config.source);
    const mapSheet = ss.getSheetByName(config.map);

    validateRequirements(sourceSheet, mapSheet, config);

    const sourceRange = sourceSheet.getDataRange();
    const sourceData = sourceRange.getDisplayValues();
    const sourceHeaders = sourceData.shift();

    if (sourceHeaders.length === 0)
      throw new Error("La hoja de origen no tiene encabezados.");

    const headerIndex = {};
    sourceHeaders.forEach((h, i) => (headerIndex[h] = i));

    const { filterRules, outputColumns } = parseRules(mapSheet, sourceHeaders);
    if (outputColumns.length === 0)
      throw new Error("No se encontraron columnas de salida válidas.");

    const sandboxSheet =
      ss.getSheetByName("_EVAL_SANDBOX_") || ss.insertSheet("_EVAL_SANDBOX_");
    sandboxSheet.hideSheet();

    const finalData = [];

    sourceData.forEach((row) => {
      const outputRowRefs = {};
      const outputValuesForFilter = {};
      const currentRowNum = finalData.length + 2;

      const processedRow = outputColumns.map((col, idx) => {
        let val = col.instruction;

        sourceHeaders.forEach((h, i) => {
          const pattern = `${PREFIX.SRC}${SYMBOLS.OPEN}${h}${SYMBOLS.CLOSE}`;
          if (val.includes(pattern)) {
            let replacement = row[i];
            if (col.isFormula) {
              replacement = formatForFormula(replacement);
            }
            val = val.split(pattern).join(replacement);
          }
        });

        Object.keys(outputRowRefs).forEach((h) => {
          const pattern = `${PREFIX.SELF}${SYMBOLS.OPEN}${h}${SYMBOLS.CLOSE}`;
          if (val.includes(pattern))
            val = val.split(pattern).join(outputRowRefs[h]);
        });

        let res;
        if (col.type === "CONSTANT" || col.type === "FORMULA") {
          res = val;
        } else {
          const srcIdx = headerIndex[val];
          res = srcIdx !== undefined ? row[srcIdx] : val;
        }

        outputRowRefs[col.header] = getColumnLetter(idx + 1) + currentRowNum;
        outputValuesForFilter[col.header] = res;
        return res;
      });

      let passes = true;
      for (const f of filterRules) {
        if (!f.isEval) continue;

        let cond = f.instruction;
        sourceHeaders.forEach((h, i) => {
          const pattern = `${PREFIX.SRC}${SYMBOLS.OPEN}${h}${SYMBOLS.CLOSE}`;
          if (cond.includes(pattern)) {
            let replacement = row[i];
            if (f.instruction.includes(PREFIX.FORMULA)) {
              replacement = formatForFormula(replacement);
            }
            cond = cond.split(pattern).join(replacement);
          }
        });

        Object.keys(outputValuesForFilter).forEach((h) => {
          const pattern = `${PREFIX.SELF}${SYMBOLS.OPEN}${h}${SYMBOLS.CLOSE}`;
          if (cond.includes(pattern))
            cond = cond.split(pattern).join(outputValuesForFilter[h]);
        });

        const filterResult = safeEval(cond, f.header, sandboxSheet);
        if (!filterResult) {
          passes = false;
          break;
        }
      }

      if (!passes) return;
      finalData.push(processedRow);
    });

    const outputHeaders = outputColumns.map((c) => c.header);
    finalData.unshift(outputHeaders);

    const outputSheet = ss.insertSheet(config.output);
    outputSheet
      .getRange(1, 1, finalData.length, outputHeaders.length)
      .setValues(finalData);

    const resultRows = finalData.length - 1;

    const summary = [
      `📊 Filas: ${resultRows}`,
      `📥 Src: ${config.source}`,
      `📑 Out: ${config.output}`,
    ].join("  |  ");

    SpreadsheetApp.getActiveSpreadsheet().toast(
      summary,
      "ETL Engine - Operación Exitosa!",
      10,
    );
    return { rows: resultRows };
  } catch (error) {
    console.error("Stack:", error.stack);
    ui.alert("Error de ejecución", error.message, ui.ButtonSet.OK);
    throw error;
  }
}

// --- Support functions ---

function validateRequirements(source, map, config) {
  if (!source || !map)
    throw new Error("Hojas de origen o mapeo no encontradas.");
}

function validateDelimiters(text, header) {
  let count = 0;
  for (let char of text) {
    if (char === SYMBOLS.OPEN) count++;
    if (char === SYMBOLS.CLOSE) count--;
    if (count < 0)
      throw new Error(
        `Llave de cierre "${SYMBOLS.CLOSE}" extra en regla "${header}".`,
      );
  }
  if (count !== 0)
    throw new Error(`Llave "${SYMBOLS.OPEN}" sin cerrar en regla "${header}".`);
}

function parseRules(mapSheet, sourceHeaders) {
  const rawRules = mapSheet
    .getDataRange()
    .getValues()
    .slice(1)
    .filter(
      (r) => r[0] && !r[0].toString().trim().startsWith(PREFIX.COMMENT_LINE),
    );

  const filterRules = [];
  const outputColumns = [];

  rawRules.forEach((r) => {
    const header = r[0].toString().trim();
    const rawInstruction = r[1].toString().trim();

    validateDelimiters(rawInstruction, header);

    const escapedOpen = SYMBOLS.OPEN.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const escapedClose = SYMBOLS.CLOSE.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const srcRegex = new RegExp(
      `${PREFIX.SRC}${escapedOpen}([^${escapedClose}]+)${escapedClose}`,
      "g",
    );

    let match;
    while ((match = srcRegex.exec(rawInstruction)) !== null) {
      const colName = match[1];
      if (!sourceHeaders.includes(colName)) {
        throw new Error(
          `Columna "${colName}" no existe en origen (regla: "${header}").`,
        );
      }
    }

    if (header.startsWith(PREFIX.FILTER)) {
      let instruction = rawInstruction.replace(PREFIX.EVAL, "").trim();

      if (instruction.startsWith(PREFIX.FORMULA)) {
        const formulaBody = instruction.replace(PREFIX.FORMULA, "").trim();
        if (!formulaBody.startsWith("=")) {
          throw new Error(
            `Error en "${header}": Después de "${PREFIX.FORMULA}" debe seguir el símbolo "=" inicial.`,
          );
        }
      }

      filterRules.push({
        header: header,
        isEval: rawInstruction.startsWith(PREFIX.EVAL),
        instruction: instruction,
      });
    } else {
      let type = "DIRECT";
      let instruction = rawInstruction;

      if (rawInstruction.startsWith(PREFIX.CONSTANT)) {
        type = "CONSTANT";
        instruction = rawInstruction.replace(PREFIX.CONSTANT, "").trim();
      } else if (rawInstruction.startsWith(PREFIX.FORMULA)) {
        type = "FORMULA";
        const formulaBody = rawInstruction.replace(PREFIX.FORMULA, "").trim();

        if (!formulaBody.startsWith("=")) {
          throw new Error(
            `Error en "${header}": Después de "${PREFIX.FORMULA}" debe seguir el símbolo "=" inicial.`,
          );
        }

        instruction = formulaBody;
      }

      outputColumns.push({
        header,
        instruction,
        type,
        isFormula: type === "FORMULA",
      });
    }
  });

  return { filterRules, outputColumns };
}

function formatForFormula(value) {
  if (value === null || value === undefined || value === "") return '""';
  const str = value.toString().trim();
  const dateMatch = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (dateMatch) {
    return `DATE(${dateMatch[3]};${dateMatch[2]};${dateMatch[1]})`;
  }
  return isNaN(str.replace("%", "")) ? `"${str}"` : str;
}

function safeEval(expression, contextHeader, sandboxSheet) {
  const operatorsLogic = {
    [SYMBOLS.OP_EQUAL]: (a, b) => a == b,
    [SYMBOLS.OP_NOT_EQUAL]: (a, b) => a != b,
    [SYMBOLS.OP_GREATER_EQUAL]: (a, b) => parseFloat(a) >= parseFloat(b),
    [SYMBOLS.OP_LESS_EQUAL]: (a, b) => parseFloat(a) <= parseFloat(b),
    [SYMBOLS.OP_GREATER]: (a, b) => parseFloat(a) > parseFloat(b),
    [SYMBOLS.OP_LESS]: (a, b) => parseFloat(a) < parseFloat(b),
  };

  const conditions = expression.split(SYMBOLS.OR);

  return conditions.some((cond) => {
    const trimmed = cond.trim();

    if (trimmed.startsWith(PREFIX.FORMULA)) {
      const formulaPart = trimmed.replace(PREFIX.FORMULA, "").trim();
      const cell = sandboxSheet.getRange("A1");
      cell.clear();
      cell.setFormula(formulaPart);
      SpreadsheetApp.flush();
      const result = cell.getValue();
      const isTrue =
        result === true || result.toString().toUpperCase() === "TRUE";

      console.log(
        `[DEBUG-FORMULA] ${contextHeader}: ${formulaPart} => ${result} (Verdict: ${isTrue})`,
      );
      return isTrue;
    }

    const op = [
      SYMBOLS.OP_EQUAL,
      SYMBOLS.OP_NOT_EQUAL,
      SYMBOLS.OP_GREATER_EQUAL,
      SYMBOLS.OP_LESS_EQUAL,
      SYMBOLS.OP_GREATER,
      SYMBOLS.OP_LESS,
    ].find((o) => trimmed.includes(o));

    if (!op) return false;

    let parts = trimmed.split(op).map((p) => p.trim());

    parts = parts.map((p) => {
      if (p.startsWith("=")) {
        sandboxSheet.getRange("A1").setFormula(p);
        SpreadsheetApp.flush();
        return sandboxSheet.getRange("A1").getValue();
      }
      return p.replace(/^"|"$/g, "");
    });

    const result = operatorsLogic[op](parts[0], parts[1]);
    console.log(
      `[DEBUG] ${contextHeader}: ${parts[0]} ${op} ${parts[1]} => ${result}`,
    );
    return result;
  });
}

function getColumnLetter(col) {
  let letter = "";
  while (col > 0) {
    let temp = (col - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    col = Math.floor((col - temp) / 26);
  }
  return letter;
}

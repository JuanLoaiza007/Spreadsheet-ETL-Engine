/**
 * @fileoverview Spreadsheet ETL Engine
 * @version 1.0.2
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
function runMapping() {
  const ui = SpreadsheetApp.getUi();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = loadConfig(ss);
    validateConfigNames(config);

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

    const finalData = [];

    sourceData.forEach((row) => {
      const outputRowRefs = {};
      const currentRowNum = finalData.length + 2;

      const passes = filterRules.every((f) => {
        if (!f.isEval) return true;
        let cond = f.instruction;
        sourceHeaders.forEach((h, i) => {
          const pattern = `${PREFIX.SRC}${SYMBOLS.OPEN}${h}${SYMBOLS.CLOSE}`;
          if (cond.includes(pattern)) cond = cond.split(pattern).join(row[i]);
        });
        return safeEval(cond, f.header);
      });

      if (!passes) return;

      const processedRow = outputColumns.map((col, idx) => {
        let val = col.instruction;

        sourceHeaders.forEach((h, i) => {
          const pattern = `${PREFIX.SRC}${SYMBOLS.OPEN}${h}${SYMBOLS.CLOSE}`;
          if (val.includes(pattern)) {
            let replacement = row[i];
            if (
              col.isFormula &&
              isNaN(replacement.toString().replace("%", ""))
            ) {
              replacement = `"${replacement}"`;
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
        return res;
      });

      finalData.push(processedRow);
    });

    const outputHeaders = outputColumns.map((c) => c.header);
    finalData.unshift(outputHeaders);

    const outputSheet =
      ss.getSheetByName(config.output) || ss.insertSheet(config.output);
    outputSheet
      .clear()
      .getRange(1, 1, finalData.length, outputHeaders.length)
      .setValues(finalData);

    ui.alert(
      "Éxito",
      `Proceso completado. Se generaron ${finalData.length - 1} filas.`,
      ui.ButtonSet.OK,
    );
  } catch (error) {
    console.error("Stack:", error.stack);
    ui.alert("Error de ejecución", error.message, ui.ButtonSet.OK);
  }
}

// --- Support functions ---

function loadConfig(ss) {
  const dashSheet = ss.getSheetByName("Dashboard");
  let cfg = { source: "Source", map: "Map", output: "Output" };
  if (dashSheet) {
    dashSheet
      .getDataRange()
      .getValues()
      .forEach((row) => {
        if (!row[0]) return;
        const key = row[0].toString().toLowerCase().trim();
        if (cfg.hasOwnProperty(key)) cfg[key] = row[1].toString().trim();
      });
  }
  return cfg;
}

function validateConfigNames(config) {
  ["source", "map", "output"].forEach((key) => {
    if (!config[key] || config[key].trim() === "")
      throw new Error(`Falta nombre de hoja "${key}".`);
  });
}

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
      filterRules.push({
        header: header,
        isEval: rawInstruction.startsWith(PREFIX.EVAL),
        instruction: rawInstruction.replace(PREFIX.EVAL, "").trim(),
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

        const equalsCount = (formulaBody.match(/=/g) || []).length;
        if (!formulaBody.startsWith("=") || equalsCount !== 1) {
          throw new Error(
            `Error en "${header}": Después de "${PREFIX.FORMULA}" debe seguir exactamente un símbolo "=" inicial para la fórmula.`,
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

function safeEval(expression, contextHeader = "Filtro") {
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

    const invalidOpMatch = trimmed.match(/[=><!]{3,}|[><]{2,}/);
    if (invalidOpMatch) {
      throw new Error(
        `Operador "${invalidOpMatch[0]}" inválido en "${contextHeader}".`,
      );
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

    const parts = trimmed.split(op).map((p) => p.trim().replace(/^"|"$/g, ""));

    if (parts.length !== 2) return false;
    return operatorsLogic[op](parts[0], parts[1]);
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

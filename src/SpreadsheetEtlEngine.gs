/**
 * @fileoverview Spreadsheet ETL Engine
 * @version 0.0.2-alpha
 * @author JuanLoaiza007
 * @license MIT
 * @description A dynamic engine for extracting, transforming, and loading data within
 * Google Sheets. It features a custom expressive language for rule-based processing,
 * built-in syntax validation, and automated error handling.
 */

const DEFAULT_CONFIG = {
  source: "Source",
  map: "Map",
  output: "Output",
};

// =============================================================================
// LEXER - Tokenizador
// Responsabilidad: Convertir texto en tokens identificables
// =============================================================================

const TokenType = {
  SRC_REF: "SRC_REF", // src[ColumnName]
  SELF_REF: "SELF_REF", // self[ColumnName]
  STRING: "STRING", // "hello"
  NUMBER: "NUMBER", // 42
  OPERATOR: "OPERATOR", // == != > < >= <=
  OR: "OR", // ||
  FORMULA: "FORMULA", // =SUM(...)
  TEXT: "TEXT", // texto literal
};

const Lexer = {
  // Patrones de reconocimiento en orden de precedencia
  // IMPORTANTE: OPERATOR debe ir antes de FORMULA para que "==" no sea capturado como fórmula
  patterns: [
    { type: TokenType.SRC_REF, regex: /src\[([^\]]+)\]/g },
    { type: TokenType.SELF_REF, regex: /self\[([^\]]+)\]/g },
    { type: TokenType.OPERATOR, regex: /(==|!=|>=|<=|>|<)/g },
    { type: TokenType.FORMULA, regex: /(=[^"'\s\|\]=]+)/g },
    { type: TokenType.STRING, regex: /"([^"]*)"/g },
    // Números con "." o "," como separador decimal (formato internacional)
    { type: TokenType.NUMBER, regex: /(\d+[.,]\d+|\d+)/g },
    { type: TokenType.OR, regex: /\|\|/g },
  ],

  /**
   * Tokeniza una cadena de entrada
   * @param {string} input - Cadena a tokenizar
   * @returns {Array<{type, value, raw}>} Array de tokens
   */
  tokenize(input) {
    const tokens = [];
    let remaining = input.trim();

    while (remaining.length > 0) {
      remaining = remaining.trimStart();
      if (remaining.length === 0) break;

      let matched = false;

      for (const { type, regex } of this.patterns) {
        regex.lastIndex = 0;
        const match = regex.exec(remaining);

        // Solo aceptar match al inicio del string
        if (match && match.index === 0) {
          tokens.push({
            type,
            value: match[1] !== undefined ? match[1] : match[0],
            raw: match[0],
          });
          remaining = remaining.slice(match[0].length);
          matched = true;
          break;
        }
      }

      // Si no hay match, consumir como texto literal hasta el siguiente token
      if (!matched) {
        // Buscar el siguiente carácter que podría iniciar un token
        const nextToken = remaining.search(
          /(src\[|self\[|"|\d|==|!=|>=|<=|>|<|\|\||=)/,
        );
        let text;

        if (nextToken === -1) {
          text = remaining;
          remaining = "";
        } else if (nextToken === 0) {
          // Si estamos al inicio y no match, tomar un carácter
          text = remaining[0];
          remaining = remaining.slice(1);
        } else {
          text = remaining.slice(0, nextToken);
          remaining = remaining.slice(nextToken);
        }

        if (text.trim()) {
          tokens.push({ type: TokenType.TEXT, value: text.trim(), raw: text });
        }
      }
    }

    return tokens;
  },
};

// =============================================================================
// PARSER - Analizador Sintáctico
// Responsabilidad: Convertir tokens en estructura parseada
// =============================================================================

const InstructionType = {
  DIRECT: "DIRECT", // Referencia directa a columna
  CONSTANT: "CONSTANT", // Valor estático
  FORMULA: "FORMULA", // Fórmula de spreadsheet
  FILTER: "FILTER", // Regla de filtro
};

const Parser = {
  /**
   * Parsea una instrucción completa
   * @param {string} header - Encabezado de la regla
   * @param {string} rawInstruction - Instrucción sin procesar
   * @param {Array<string>} sourceHeaders - Nombres de columnas fuente
   * @returns {Object} Instrucción parseada
   */
  parseInstruction(header, rawInstruction, sourceHeaders) {
    const instruction = rawInstruction.trim();

    // Validar delimitadores balanceados
    this.validateDelimiters(instruction, header);

    // Detectar tipo de instrucción por prefijo
    if (header.startsWith("_filter:")) {
      return this.parseFilter(header, instruction, sourceHeaders);
    }

    if (instruction.startsWith("constant:")) {
      return {
        type: InstructionType.CONSTANT,
        header,
        value: instruction.replace("constant:", "").trim(),
        isFormula: false,
      };
    }

    if (instruction.startsWith("formula:")) {
      const formula = instruction.replace("formula:", "").trim();
      if (!formula.startsWith("=")) {
        throw new Error(
          `Error en "${header}": Después de "formula:" debe seguir el símbolo "=" inicial.`,
        );
      }
      return {
        type: InstructionType.FORMULA,
        header,
        formula,
        isFormula: true,
      };
    }

    // Verificar si contiene referencias o es referencia directa
    const tokens = Lexer.tokenize(instruction);
    const hasRefs = tokens.some(
      (t) => t.type === TokenType.SRC_REF || t.type === TokenType.SELF_REF,
    );

    // Si es un solo token de texto, podría ser referencia directa a columna
    if (tokens.length === 1 && tokens[0].type === TokenType.TEXT) {
      return {
        type: InstructionType.DIRECT,
        header,
        column: tokens[0].value,
        isFormula: false,
      };
    }

    // Es una expresión con tokens mezclados
    return {
      type: "EXPRESSION",
      header,
      tokens,
      raw: instruction,
      isFormula: hasRefs,
    };
  },

  /**
   * Parsea una regla de filtro
   */
  parseFilter(header, instruction, sourceHeaders) {
    const isEval = instruction.startsWith("eval:");
    const expr = isEval ? instruction.replace("eval:", "").trim() : instruction;

    // Detectar si es fórmula completa
    if (expr.startsWith("formula:")) {
      const formula = expr.replace("formula:", "").trim();
      if (!formula.startsWith("=")) {
        throw new Error(
          `Error en "${header}": Después de "formula:" debe seguir el símbolo "=" inicial.`,
        );
      }
      return {
        type: InstructionType.FILTER,
        header,
        isEval,
        isFormula: true,
        formula,
      };
    }

    // Validar referencias a columnas fuente
    this.validateSourceRefs(expr, header, sourceHeaders);

    // Parsear expresión con condiciones OR
    const conditions = this.parseConditions(expr);
    return {
      type: InstructionType.FILTER,
      header,
      isEval,
      isFormula: false,
      conditions,
    };
  },

  /**
   * Parsea condiciones separadas por ||
   */
  parseConditions(expression) {
    const parts = expression.split("||").map((p) => p.trim());

    return parts.map((part) => {
      const tokens = Lexer.tokenize(part);
      const opToken = tokens.find((t) => t.type === TokenType.OPERATOR);

      // Si no hay operador, podría ser una fórmula
      if (!opToken) {
        const formulaToken = tokens.find((t) => t.type === TokenType.FORMULA);
        if (formulaToken) {
          return {
            left: null,
            op: "FORMULA",
            right: { type: "FORMULA", formula: formulaToken.value },
          };
        }
        throw new Error(`Condición sin operador válido: "${part}"`);
      }

      const opIndex = tokens.indexOf(opToken);
      const left = this.parseOperand(tokens.slice(0, opIndex));
      const right = this.parseOperand(tokens.slice(opIndex + 1));

      return { left, op: opToken.value, right };
    });
  },

  /**
   * Parsea un operando (lado izquierdo o derecho de una condición)
   */
  parseOperand(tokens) {
    if (tokens.length === 0) return null;

    // Concatenar tokens si son múltiples
    if (tokens.length > 1) {
      return {
        type: "EXPRESSION",
        tokens,
        raw: tokens.map((t) => t.raw).join(" "),
      };
    }

    const token = tokens[0];

    switch (token.type) {
      case TokenType.SRC_REF:
        return { type: TokenType.SRC_REF, column: token.value };
      case TokenType.SELF_REF:
        return { type: TokenType.SELF_REF, column: token.value };
      case TokenType.STRING:
        return { type: TokenType.STRING, value: token.value };
      case TokenType.NUMBER:
        // Convertir "," a "." para parseFloat
        const numValue = parseFloat(token.value.replace(",", "."));
        return { type: TokenType.NUMBER, value: numValue };
      case TokenType.FORMULA:
        return { type: TokenType.FORMULA, formula: token.value };
      default:
        return { type: TokenType.TEXT, value: token.value };
    }
  },

  /**
   * Valida que los delimitadores [ ] estén balanceados
   */
  validateDelimiters(text, header) {
    let count = 0;
    for (const char of text) {
      if (char === "[") count++;
      if (char === "]") count--;
      if (count < 0) {
        throw new Error(`Llave de cierre "]" extra en regla "${header}".`);
      }
    }
    if (count !== 0) {
      throw new Error(`Llave "[" sin cerrar en regla "${header}".`);
    }
  },

  /**
   * Valida que las referencias src[...] apunten a columnas existentes
   */
  validateSourceRefs(instruction, header, sourceHeaders) {
    const regex = /src\[([^\]]+)\]/g;
    let match;
    while ((match = regex.exec(instruction)) !== null) {
      if (!sourceHeaders.includes(match[1])) {
        throw new Error(
          `Columna "${match[1]}" no existe en origen (regla: "${header}").`,
        );
      }
    }
  },
};

// =============================================================================
// EVALUATOR - Evaluador
// Responsabilidad: Ejecutar la lógica de evaluación sobre datos
// =============================================================================

const Evaluator = {
  operators: {
    "==": (a, b) => a == b,
    "!=": (a, b) => a != b,
    ">": (a, b) => parseFloat(a) > parseFloat(b),
    "<": (a, b) => parseFloat(a) < parseFloat(b),
    ">=": (a, b) => parseFloat(a) >= parseFloat(b),
    "<=": (a, b) => parseFloat(a) <= parseFloat(b),
  },

  /**
   * Evalúa un filtro completo contra una fila de datos
   * @param {Object} filter - Filtro parseado
   * @param {Object} context - Contexto con sourceData y outputData
   * @param {Sheet} sandbox - Hoja sandbox para evaluar fórmulas
   * @returns {boolean} true si la fila pasa el filtro
   */
  evaluateFilter(filter, context, sandbox) {
    if (!filter.isEval) return true;

    // Filtro basado en fórmula
    if (filter.isFormula) {
      let formula = this.resolveTokens(filter.formula, context);
      return this.evaluateFormula(formula, sandbox, filter.header);
    }

    // Filtro basado en condiciones OR
    return filter.conditions.some((cond) => {
      if (cond.op === "FORMULA") {
        let formula = this.resolveTokens(cond.right.formula, context);
        return this.evaluateFormula(formula, sandbox, filter.header);
      }
      return this.evaluateCondition(cond, context, sandbox);
    });
  },

  /**
   * Evalúa una condición individual
   */
  evaluateCondition(condition, context, sandbox) {
    const left = this.resolveOperand(condition.left, context, sandbox);
    const right = this.resolveOperand(condition.right, context, sandbox);
    const result = this.operators[condition.op](left, right);

    console.log(
      `[DEBUG-COND] ${context.header}: ${JSON.stringify(condition.left?.type || "raw")}:${left} ${condition.op} ${JSON.stringify(condition.right?.type || "raw")}:${right} => ${result}`,
    );
    return result;
  },

  /**
   * Resuelve el valor de un operando
   */
  resolveOperand(operand, context, sandbox) {
    if (!operand) return "";

    switch (operand.type) {
      case TokenType.SRC_REF:
        const srcVal = context.sourceData[operand.column] || "";
        console.log(`[RESOLVE] SRC_REF[${operand.column}] = "${srcVal}"`);
        return srcVal;
      case TokenType.SELF_REF:
        const selfVal = context.outputData[operand.column] || "";
        console.log(
          `[RESOLVE] SELF_REF[${operand.column}] = "${selfVal}" (outputData: ${JSON.stringify(context.outputData)})`,
        );
        return selfVal;
      case TokenType.FORMULA:
        return this.evaluateFormula(operand.formula, sandbox, context.header);
      case TokenType.STRING:
      case TokenType.NUMBER:
        return operand.value;
      case "EXPRESSION":
        return this.resolveTokens(operand.raw, context);
      default:
        return operand.value || "";
    }
  },

  /**
   * Resuelve todos los tokens src[...] y self[...] en un string
   */
  resolveTokens(template, context) {
    let result = template;

    // Resolver src[...]
    const srcRegex = /src\[([^\]]+)\]/g;
    result = result.replace(srcRegex, (_, col) => {
      const val = context.sourceData[col] || "";
      return context.isFormula ? formatForFormula(val) : val;
    });

    // Resolver self[...]
    const selfRegex = /self\[([^\]]+)\]/g;
    result = result.replace(selfRegex, (_, col) => {
      return context.outputData[col] || "";
    });

    return result;
  },

  /**
   * Evalúa una fórmula usando el sandbox
   */
  evaluateFormula(formula, sandbox, contextHeader) {
    const cell = sandbox.getRange("A1");
    cell.clear();
    cell.setFormula(formula);
    SpreadsheetApp.flush();
    const result = cell.getValue();
    const isTrue =
      result === true || result.toString().toUpperCase() === "TRUE";

    console.log(
      `[DEBUG-FORMULA] ${contextHeader}: ${formula} => ${result} (Verdict: ${isTrue})`,
    );
    return isTrue;
  },
};

// =============================================================================
// HELPERS - Funciones auxiliares
// =============================================================================

/**
 * Formatea un valor para uso dentro de fórmulas de spreadsheet
 */
function formatForFormula(value) {
  if (value === null || value === undefined || value === "") return '""';
  const str = value.toString().trim();

  // Detectar formato de fecha DD/MM/YYYY o DD-MM-YYYY
  const dateMatch = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (dateMatch) {
    return `DATE(${dateMatch[3]};${dateMatch[2]};${dateMatch[1]})`;
  }

  // Normalizar separador decimal: "," -> "."
  const normalizedStr = str.replace(",", ".");

  // Si es texto, entrecomillar; si es número, dejarlo
  return isNaN(normalizedStr.replace("%", "")) ? `"${str}"` : normalizedStr;
}

/**
 * Convierte índice de columna a letra (1=A, 2=B, ..., 27=AA)
 */
function getColumnLetter(col) {
  let letter = "";
  while (col > 0) {
    const temp = (col - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    col = Math.floor((col - temp) / 26);
  }
  return letter;
}

// =============================================================================
// MAIN - Función principal de ejecución
// =============================================================================

/**
 * Main function that executes the transformation process.
 */
function runMapping(config = DEFAULT_CONFIG) {
  const ui = SpreadsheetApp.getUi();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(config.source);
    const mapSheet = ss.getSheetByName(config.map);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Iniciando mapeo, espera un momento.`,
      "ETL Engine",
      4,
    );

    if (!sourceSheet || !mapSheet) {
      throw new Error("Hojas de origen o mapeo no encontradas.");
    }

    const sourceRange = sourceSheet.getDataRange();
    const sourceData = sourceRange.getDisplayValues();
    const sourceHeaders = sourceData.shift();

    if (sourceHeaders.length === 0) {
      throw new Error("La hoja de origen no tiene encabezados.");
    }

    const headerIndex = {};
    sourceHeaders.forEach((h, i) => (headerIndex[h] = i));

    // Parsear reglas usando el nuevo Parser
    const { filterRules, outputColumns } = parseRules(mapSheet, sourceHeaders);
    if (outputColumns.length === 0) {
      throw new Error("No se encontraron columnas de salida válidas.");
    }

    // Crear sandbox para evaluación de fórmulas
    const sandboxSheet =
      ss.getSheetByName("_EVAL_SANDBOX_") || ss.insertSheet("_EVAL_SANDBOX_");
    sandboxSheet.hideSheet();

    const finalData = [];

    // Procesar cada fila de datos
    sourceData.forEach((row) => {
      const outputRowRefs = {};
      const outputValuesForFilter = {};
      const currentRowNum = finalData.length + 2;

      // Procesar columnas de salida
      const processedRow = outputColumns.map((col, idx) => {
        const value = processInstruction(
          col,
          row,
          headerIndex,
          outputRowRefs,
          outputValuesForFilter,
        );

        // Registrar referencia de columna para self[...]
        outputRowRefs[col.header] = getColumnLetter(idx + 1) + currentRowNum;

        // Para filtros: si es fórmula, evaluar el valor real
        console.log(
          `[DEBUG-FORMULA-CHECK] ${col.header}: isFormula=${col.isFormula}, type=${col.type}, valueStartsWithEq=${value.toString().startsWith("=")}`,
        );

        if (col.isFormula && value.toString().startsWith("=")) {
          const cell = sandboxSheet.getRange("B1");
          cell.setFormula(value);
          SpreadsheetApp.flush();
          const evaluatedValue = cell.getValue();
          outputValuesForFilter[col.header] = evaluatedValue;
          console.log(
            `[EVAL-OUTPUT] ${col.header}: ${value} => ${evaluatedValue}`,
          );
        } else {
          outputValuesForFilter[col.header] = value;
        }

        return value;
      });

      // Evaluar filtros
      const context = {
        sourceData: Object.fromEntries(
          sourceHeaders.map((h, i) => [h, row[i]]),
        ),
        outputData: outputValuesForFilter,
        header: "filter",
        isFormula: false,
      };

      let passes = true;
      let failedFilter = null;
      let filterResult = null;

      for (const filter of filterRules) {
        context.header = filter.header;
        context.isFormula = filter.isFormula;

        const result = Evaluator.evaluateFilter(filter, context, sandboxSheet);
        if (!result) {
          passes = false;
          failedFilter = filter.header;
          filterResult = result;
          break;
        }
      }

      if (passes) {
        finalData.push(processedRow);
      } else {
        // Mensaje de depuración para filas descartadas
        const rowIdentifier = row[0] || `Fila ${sourceData.indexOf(row) + 2}`;
        console.log(
          `[FILTER-DISCARD] Fila "${rowIdentifier}" descartada por filtro "${failedFilter}" (resultado: ${filterResult})`,
        );
      }
    });

    // Preparar datos de salida
    const outputHeaders = outputColumns.map((c) => c.header);
    finalData.unshift(outputHeaders);

    // Escribir resultados
    const outputSheet =
      ss.getSheetByName(config.output) || ss.insertSheet(config.output);
    outputSheet
      .clear()
      .getRange(1, 1, finalData.length, outputHeaders.length)
      .setValues(finalData);

    const summary = [
      `📊 Filas: ${finalData.length - 1}`,
      `📥 Src: ${config.source}`,
      `📑 Out: ${config.output}`,
    ].join("  |  ");

    SpreadsheetApp.getActiveSpreadsheet().toast(
      summary,
      "ETL Engine - Operación Exitosa!",
      10,
    );
  } catch (error) {
    console.error("Stack:", error.stack);
    ui.alert("Error de ejecución", error.message, ui.ButtonSet.OK);
  }
}

/**
 * Parsea las reglas del mapa usando el Parser estructurado
 */
function parseRules(mapSheet, sourceHeaders) {
  const rawRules = mapSheet
    .getDataRange()
    .getValues()
    .slice(1)
    .filter((r) => r[0] && !r[0].toString().trim().startsWith("//"));

  const filterRules = [];
  const outputColumns = [];

  rawRules.forEach((r) => {
    const header = r[0].toString().trim();
    const rawInstruction = r[1].toString().trim();

    const parsed = Parser.parseInstruction(
      header,
      rawInstruction,
      sourceHeaders,
    );

    if (parsed.type === InstructionType.FILTER) {
      filterRules.push(parsed);
    } else {
      outputColumns.push(parsed);
    }
  });

  return { filterRules, outputColumns };
}

/**
 * Procesa una instrucción y devuelve el valor resultante
 */
function processInstruction(
  instruction,
  row,
  headerIndex,
  outputRowRefs,
  outputValuesForFilter,
) {
  let result;

  switch (instruction.type) {
    case InstructionType.CONSTANT:
      result = instruction.value;
      break;

    case InstructionType.FORMULA:
      result = resolveInstruction(
        instruction.formula,
        row,
        headerIndex,
        outputRowRefs,
        true,
      );
      break;

    case InstructionType.DIRECT:
      const srcIdx = headerIndex[instruction.column];
      result = srcIdx !== undefined ? row[srcIdx] : instruction.column;
      break;

    case "EXPRESSION":
      result = resolveInstruction(
        instruction.raw,
        row,
        headerIndex,
        outputRowRefs,
        instruction.isFormula,
      );
      break;

    default:
      result = instruction.raw || instruction.value || "";
  }

  console.log(
    `[PROCESS] ${instruction.header} (${instruction.type}) => "${result}"`,
  );
  return result;
}

/**
 * Resuelve una instrucción con tokens src[...] y self[...]
 */
function resolveInstruction(
  template,
  row,
  headerIndex,
  outputRowRefs,
  isFormula,
) {
  let result = template;

  // Resolver src[...]
  const srcRegex = /src\[([^\]]+)\]/g;
  result = result.replace(srcRegex, (_, col) => {
    const idx = headerIndex[col];
    const val = idx !== undefined ? row[idx] : "";
    return isFormula ? formatForFormula(val) : val;
  });

  // Resolver self[...]
  const selfRegex = /self\[([^\]]+)\]/g;
  result = result.replace(selfRegex, (_, col) => {
    return outputRowRefs[col] || "";
  });

  return result;
}

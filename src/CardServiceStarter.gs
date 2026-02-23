/**
 * @fileoverview
 * @version 1.1.0
 * @author JuanLoaiza007
 * @license MIT
 * @description Volatile UI for the Spreadsheet ETL Engine
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Utilidades")
    .addItem("Ejecutar ETL", "executeETL")
    .addToUi();
}

/** --- UI AND EVENTS --- **/

function executeETL() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  const sheetList = sheets.map((s, i) => `${i + 1}. ${s.getName()}`).join("\n");

  const res = ui.prompt(
    "Ejecutar ETL",
    `Selecciona las hojas (Source, Map) usando sus números:\n\n${sheetList}\n\nFormato: int, int (ej: 1, 2)`,
    ui.ButtonSet.OK_CANCEL,
  );

  if (res.getSelectedButton() !== ui.Button.OK) return;

  const input = res
    .getResponseText()
    .split(",")
    .map((s) => parseInt(s.trim(), 10) - 1);

  if (input.length !== 2 || isNaN(input[0]) || isNaN(input[1])) {
    ui.alert(
      "Error",
      "Formato inválido. Debe ser 'int, int'.",
      ui.ButtonSet.OK,
    );
    return;
  }

  const sourceName = sheets[input[0]] ? sheets[input[0]].getName() : null;
  const mapName = sheets[input[1]] ? sheets[input[1]].getName() : null;

  if (!sourceName || !mapName) {
    ui.alert("Error", "Números de hoja fuera de rango.", ui.ButtonSet.OK);
    return;
  }

  const timestamp = Utilities.formatDate(
    new Date(),
    ss.getSpreadsheetTimeZone(),
    "yyyy-MM-dd HH:mm:ss",
  );

  const config = {
    source: sourceName,
    map: mapName,
    output: `OUTPUT ${timestamp}`,
  };

  try {
    runMapping(config);
  } catch (e) {
    ui.alert("Error", e.message, ui.ButtonSet.OK);
  }
}
